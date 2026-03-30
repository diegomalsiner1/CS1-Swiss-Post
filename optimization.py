from ortools.linear_solver import pywraplp
import pandas as pd
from tqdm import tqdm

## Will be used to define the optimization variables, constraints, and the objective function


def setup(
    input_dict: dict,
    debug_infeasibility: bool = False,
    fixed_battery_capacity_kwh: float | None = None,
):
    """Creates and returns an OR-Tools optimization model and handles.

    Parameters:
      input_dict: problem data including time series and parameters
      debug_infeasibility: if True, add slack variables to diagnose infeasibility.

    Returns:
      model, slacks_list, solution_handles
    """

    print("Setting up optimization problem")
    model = pywraplp.Solver.CreateSolver("CBC")  # MILP solver
    if model is None:
        raise RuntimeError(
            "Could not create CBC solver. Check OR-Tools installation and CBC backend availability."
        )

    # Basic input consistency checks
    if "total_demand" not in input_dict or "PV_capacity_factor" not in input_dict or "electricity_price" not in input_dict:
        raise ValueError("input_dict must contain total_demand, PV_capacity_factor, and electricity_price time series")
    n = len(input_dict["total_demand"])
    if not (len(input_dict["PV_capacity_factor"]) == n and len(input_dict["electricity_price"]) == n):
        raise ValueError("Time series lengths in input_dict must match")

    ## Defining the required constants
    PV_max_capacity = input_dict["parameters"]["PV_max_capacity"]
    Battery_max_inflow = input_dict["parameters"]["Battery_max_inflow"]
    Battery_max_outflow = input_dict["parameters"]["Battery_max_outflow"]
    Battery_capacity_upper_bound = input_dict["parameters"]["Battery_max_capacity"]
    eta_charge = input_dict["parameters"]["Battery_eta_charge"]
    eta_discharge = input_dict["parameters"]["Battery_eta_discharge"]
    eta_self_discharge = input_dict["parameters"]["Battery_eta_self_discharge"]
    battery_invest_cost = input_dict["parameters"]["Battery_invest_cost"]
    cost_operation_and_maintenance = input_dict["parameters"]["operation_and_maintenance"]

    interest_rate = input_dict["parameters"]["interest_rate"]
    lifetime = input_dict["parameters"]["lifetime"]
    battery_degrading = input_dict["parameters"]["battery_degrading"]
    peak_shaving_cost_factor = input_dict["parameters"].get("peak_shaving_cost_factor", 0.0)
    peak_shaving_frequency = input_dict["parameters"].get("peak_shaving_granularity", input_dict["parameters"].get("peak_shaving_frequency", "yearly")).lower()
    if peak_shaving_frequency not in {"yearly", "monthly"}:
        raise ValueError("peak_shaving_granularity must be 'yearly' or 'monthly'")
    optimization_mode = input_dict["parameters"].get("optimization_mode", "milp").lower()
    if optimization_mode not in {"milp", "lp"}:
        raise ValueError("optimization_mode must be 'milp' or 'lp'")
    surplus_handling = input_dict["parameters"].get("surplus_handling", "curtail").lower()
    if surplus_handling not in {"curtail", "must_absorb"}:
        raise ValueError("surplus_handling must be 'curtail' or 'must_absorb'")
    battery_max_c_rate = input_dict["parameters"].get("battery_max_c_rate")
    if battery_max_c_rate is not None:
        battery_max_c_rate = float(battery_max_c_rate)
        if battery_max_c_rate <= 0:
            raise ValueError("battery_max_c_rate must be positive when provided")
    min_soc_fraction = float(input_dict["parameters"].get("battery_min_soc_fraction", 0.0))
    if not 0 <= min_soc_fraction < 1:
        raise ValueError("battery_min_soc_fraction must lie in [0, 1)")
    use_binaries = optimization_mode == "milp"

    if interest_rate == 0:
        CRF = 1 / lifetime
    else:
        CRF = (((1 + interest_rate) ** lifetime) * interest_rate) / (((1 + interest_rate) ** lifetime) - 1)

    ## Defining the timeseries of PV production, energy demand, and electricity prices


    ## Defining the required Variables
    Time_dependent_variables = {}
    # number of timesteps is driven by demand series length
    timesteps = range(len(input_dict["total_demand"]))

    inf = model.infinity()
    OPEX = model.NumVar(-inf, inf, f"Operational_Cost")
    Total_Cost = model.NumVar(-inf, inf, "Total_Cost")
    if fixed_battery_capacity_kwh is not None:
        fixed_battery_capacity_kwh = float(fixed_battery_capacity_kwh)
        if fixed_battery_capacity_kwh < 0 or fixed_battery_capacity_kwh > Battery_capacity_upper_bound:
            raise ValueError("fixed_battery_capacity_kwh must lie within [0, Battery_max_capacity]")
        Battery_capacity = model.NumVar(
            fixed_battery_capacity_kwh,
            fixed_battery_capacity_kwh,
            "Battery_Capacity",
        )
    else:
        Battery_capacity = model.NumVar(0, Battery_capacity_upper_bound, "Battery_Capacity")
    
    # Peak demand variables for grid flow
    Peak_grid_flow = model.NumVar(0, inf, "Peak_Grid_Flow_Yearly") if peak_shaving_frequency == "yearly" else None
    
    slacks = []
    print("Initializing time dependent variables")
    battery_level_vars = []
    battery_in_flow_vars = []
    battery_out_flow_vars = []
    grid_flow_vars = []
    pv_out_flow_vars = []
    spill_flow_vars = []

    for t in tqdm(timesteps):
        # Integer Variables
        Time_dependent_variables[("Battery_in_flow",t)] = model.NumVar(0, Battery_max_inflow, f"Powerflow_Batter_in_{t}")
        Time_dependent_variables[("Battery_out_flow",t)] = model.NumVar(0, Battery_max_outflow, f"Powerflow_Batter_out_{t}")
        Time_dependent_variables[("Grid_flow",t)] = model.NumVar(0, inf, f"Powerflow_Grid_{t}")
        Time_dependent_variables[("Battery_level",t)] = model.NumVar(0, Battery_capacity_upper_bound, f"Battery_Level_{t}")
        Time_dependent_variables[("PV_out_flow",t)] = model.NumVar(0, PV_max_capacity, f"PV_Powerflow_out_{t}")
        Time_dependent_variables[("Spill_flow",t)] = model.NumVar(0, PV_max_capacity, f"Powerflow_Spill_{t}")

        battery_in_flow_vars.append(Time_dependent_variables[("Battery_in_flow", t)])
        battery_out_flow_vars.append(Time_dependent_variables[("Battery_out_flow", t)])
        battery_level_vars.append(Time_dependent_variables[("Battery_level", t)])
        grid_flow_vars.append(Time_dependent_variables[("Grid_flow", t)])
        pv_out_flow_vars.append(Time_dependent_variables[("PV_out_flow", t)])
        spill_flow_vars.append(Time_dependent_variables[("Spill_flow", t)])

        if use_binaries:
            Time_dependent_variables[("Binary_battery_in_flow",t)] = model.BoolVar(f"Binary_battery_in_flow_{t}")
            Time_dependent_variables[("Binary_battery_out_flow",t)] = model.BoolVar(f"Binary_battery_out_flow_{t}")

    ## Constraint functions
    # time dependent constraints
    print("Intitializing time dependent constraint functions")
    for t in tqdm(timesteps):
        # Power Flow Balance
        power_balance_expr = (
            Time_dependent_variables[("Grid_flow", t)]
            + Time_dependent_variables[("PV_out_flow", t)]
            + Time_dependent_variables[("Battery_out_flow", t)]
            - Time_dependent_variables[("Battery_in_flow", t)]
        )
        demand_t = input_dict["total_demand"][t]
        if debug_infeasibility:
            s_balance = model.NumVar(0, inf, f"slack_power_balance_{t}")
            model.Add(power_balance_expr - demand_t <= s_balance)
            model.Add(demand_t - power_balance_expr <= s_balance)
            slacks.append(s_balance)
        else:
            model.Add(power_balance_expr == demand_t)

        # PV allocation between local use and curtailment/spillage
        pv_limit_t = input_dict["PV_capacity_factor"][t] * PV_max_capacity
        if debug_infeasibility:
            s_pv = model.NumVar(0, inf, f"slack_pv_balance_{t}")
            pv_balance_expr = (
                Time_dependent_variables[("PV_out_flow", t)]
                + Time_dependent_variables[("Spill_flow", t)]
            )
            model.Add(pv_balance_expr - pv_limit_t <= s_pv)
            model.Add(pv_limit_t - pv_balance_expr <= s_pv)
            slacks.append(s_pv)
        else:
            model.Add(
                Time_dependent_variables[("PV_out_flow", t)]
                + Time_dependent_variables[("Spill_flow", t)]
                == pv_limit_t
            )
        if surplus_handling == "must_absorb":
            model.Add(Time_dependent_variables[("Spill_flow", t)] == 0)

        # MILP mode enforces explicit mutual exclusivity via binary on/off variables.
        if use_binaries:
            model.Add(Time_dependent_variables[("Binary_battery_in_flow",t)] + Time_dependent_variables[("Binary_battery_out_flow",t)] <= 1)
            model.Add(Time_dependent_variables[("Battery_in_flow",t)] <= Time_dependent_variables[("Binary_battery_in_flow",t)] * Battery_max_inflow)
            model.Add(Time_dependent_variables[("Battery_out_flow",t)] <= Time_dependent_variables[("Binary_battery_out_flow",t)] * Battery_max_outflow)
        if battery_max_c_rate is not None:
            model.Add(Time_dependent_variables[("Battery_in_flow", t)] <= battery_max_c_rate * Battery_capacity)
            model.Add(Time_dependent_variables[("Battery_out_flow", t)] <= battery_max_c_rate * Battery_capacity)
        model.Add(Time_dependent_variables[("Battery_level", t)] <= Battery_capacity)
        model.Add(Time_dependent_variables[("Battery_level", t)] >= min_soc_fraction * Battery_capacity)
        if t != timesteps[-1]:
            # Equation to calculate the battery level in the next time step t+1 based on the current time step t
            soc_next_expr = (
                Time_dependent_variables[("Battery_level", t)] * (1 - eta_self_discharge)
                + 0.25 * eta_charge * Time_dependent_variables[("Battery_in_flow", t)]
                - 0.25 * (1 / eta_discharge) * Time_dependent_variables[("Battery_out_flow", t)]
            )
            if debug_infeasibility:
                s_soc = model.NumVar(0, inf, f"slack_soc_dyn_{t}")
                model.Add(Time_dependent_variables[("Battery_level", t + 1)] - soc_next_expr <= s_soc)
                model.Add(soc_next_expr - Time_dependent_variables[("Battery_level", t + 1)] <= s_soc)
                slacks.append(s_soc)
            else:
                model.Add(Time_dependent_variables[("Battery_level", t + 1)] == soc_next_expr)



    # starting conditions
    if debug_infeasibility:
        s_soc_start = model.NumVar(0, inf, "slack_soc_start")
        s_soc_end = model.NumVar(0, inf, "slack_soc_end")
        model.Add(Time_dependent_variables[("Battery_level", 0)] - 0.5 * Battery_capacity <= s_soc_start)
        model.Add(0.5 * Battery_capacity - Time_dependent_variables[("Battery_level", 0)] <= s_soc_start)
        model.Add(Time_dependent_variables[("Battery_level", timesteps[-1])] - 0.5 * Battery_capacity <= s_soc_end)
        model.Add(0.5 * Battery_capacity - Time_dependent_variables[("Battery_level", timesteps[-1])] <= s_soc_end)
        slacks.extend([s_soc_start, s_soc_end])
    else:
        model.Add(Time_dependent_variables[("Battery_level", 0)] == 0.5 * Battery_capacity)
        model.Add(Time_dependent_variables[("Battery_level", timesteps[-1])] == 0.5 * Battery_capacity)

    # Peak grid flow constraints
    monthly_peak_vars = {}
    if peak_shaving_cost_factor > 0:
        if peak_shaving_frequency == "yearly":
            # Peak grid flow for the entire year
            for t in timesteps:
                model.Add(Peak_grid_flow >= Time_dependent_variables[("Grid_flow", t)])
        else:  # monthly
            # Create peak variables for each month
            timestamps = pd.to_datetime(input_dict["timestamps"])
            df_temp = pd.DataFrame({
                "timestamp": timestamps,
            })
            df_temp["year_month"] = df_temp["timestamp"].dt.to_period("M")
            unique_months = df_temp["year_month"].unique()
            #print(f"Number of unique months: {len(unique_months)}")
            #print(f"Unique months: {list(unique_months)}")
            
            for month in unique_months:
                month_indices = [i for i, m in enumerate(df_temp["year_month"]) if m == month]
                peak_var = model.NumVar(0, inf, f"Peak_Grid_Flow_{month}")
                monthly_peak_vars[month] = peak_var
                for idx in month_indices:
                    model.Add(peak_var >= Time_dependent_variables[("Grid_flow", idx)])

    timestep_hours = 0.25
    import_cost_expr = sum(
        Time_dependent_variables[("Grid_flow", t)] * input_dict["electricity_price"][t] * timestep_hours
        for t in timesteps
    )
    annualized_battery_cost_expr = CRF * battery_invest_cost * Battery_capacity

    # Calculate peak demand costs based on frequency
    peak_demand_cost_expr = 0
    if peak_shaving_cost_factor > 0:
        if peak_shaving_frequency == "yearly":
            peak_demand_cost_expr = peak_shaving_cost_factor * Peak_grid_flow
        else:  # monthly
            peak_demand_cost_expr = peak_shaving_cost_factor * sum(monthly_peak_vars.values())

    # Objective function
    if debug_infeasibility:
        model.Minimize(sum(slacks))
    else:
        opex_expr = cost_operation_and_maintenance + import_cost_expr
        model.Add(OPEX == opex_expr)
        model.Add(Total_Cost == annualized_battery_cost_expr + OPEX + peak_demand_cost_expr)
        model.Minimize(Total_Cost)

    solution_handles = {
        "battery_capacity": Battery_capacity,
        "opex": OPEX,
        "total_cost": Total_Cost,
        "annualized_battery_cost_expr": annualized_battery_cost_expr,
        "import_cost_expr": import_cost_expr,
        "fixed_om_cost": cost_operation_and_maintenance,
        "peak_demand_cost_expr": peak_demand_cost_expr,
        "battery_level_vars": battery_level_vars,
        "battery_in_flow_vars": battery_in_flow_vars,
        "battery_out_flow_vars": battery_out_flow_vars,
        "grid_flow_vars": grid_flow_vars,
        "pv_out_flow_vars": pv_out_flow_vars,
        "spill_flow_vars": spill_flow_vars,
        "yearly_peak_var": Peak_grid_flow,
        "monthly_peak_vars": monthly_peak_vars,
    }

    return model, slacks, solution_handles


def optimize_model(model, slacks=None, top_n=20, debug_infeasibility=False):
    """
    Method for optimizing the a constrainted model
    Input:
    ortools.linear_solver.pywraplp.solver unsolved optimization model

    Output:
    ortools.linear_solver.pywraplp.solver solved optimization model
    """
    print("Running optimization")
    if debug_infeasibility:
        model.EnableOutput()  # Enable solver output for debugging during infeasibility checks
    status = model.Solve()

    if status == model.OPTIMAL:
        print('Optimal solution found')
        if slacks:
            nonzero_slacks = [(s.name(), s.solution_value()) for s in slacks if s.solution_value() > 1e-6]
            nonzero_slacks.sort(key=lambda x: x[1], reverse=True)
            print(f"Nonzero slacks: {len(nonzero_slacks)}")
            for name, val in nonzero_slacks[:top_n]:
                print(f"  {name}: {val:.6f}")
   
    else:
        raise ValueError(f'Did not find optimal solution: {status}')
    return model


def summarize_solution(model, solution_handles):
    yearly_peak = solution_handles["yearly_peak_var"].solution_value() if solution_handles["yearly_peak_var"] else None
    monthly_peaks = {k: v.solution_value() for k, v in solution_handles["monthly_peak_vars"].items()}
    sum_monthly = sum(monthly_peaks.values()) if monthly_peaks else 0
    curtailed_energy_kwh = sum(v.solution_value() for v in solution_handles["spill_flow_vars"]) * 0.25
    return {
        "battery_capacity_kwh": solution_handles["battery_capacity"].solution_value(),
        "objective_total_cost": model.Objective().Value(),
        "opex": solution_handles["opex"].solution_value(),
        "import_cost": solution_handles["import_cost_expr"].solution_value(),
        "fixed_om_cost": solution_handles["fixed_om_cost"],
        "annualized_battery_cost": solution_handles["annualized_battery_cost_expr"].solution_value(),
        "peak_demand_cost": solution_handles["peak_demand_cost_expr"].solution_value(),
        "yearly_peak": yearly_peak,
        "monthly_peaks": monthly_peaks,
        "sum_monthly_peaks": sum_monthly,
        "curtailed_energy_kwh": curtailed_energy_kwh,
    }


def compute_no_battery_baseline(input_dict):
    """Compute the no-battery baseline cost under the configured tariff structure."""
    timestep_hours = 0.25
    if len(input_dict.get("total_demand", [])) != len(input_dict.get("PV_capacity_factor", [])):
        raise ValueError("Input timeseries lengths must match for baseline computation.")

    peak_shaving_cost_factor = float(input_dict["parameters"].get("peak_shaving_cost_factor", 0.0))
    peak_shaving_frequency = input_dict["parameters"].get(
        "peak_shaving_granularity",
        input_dict["parameters"].get("peak_shaving_frequency", "yearly"),
    ).lower()
    surplus_handling = input_dict["parameters"].get("surplus_handling", "curtail").lower()
    if peak_shaving_frequency not in {"yearly", "monthly"}:
        raise ValueError("peak_shaving_granularity must be 'yearly' or 'monthly'")

    grid_import_series = []
    for t, demand in enumerate(input_dict["total_demand"]):
        pv_limit = input_dict["PV_capacity_factor"][t] * input_dict["parameters"]["PV_max_capacity"]
        if surplus_handling == "must_absorb" and pv_limit > demand + 1e-9:
            raise ValueError(
                "No-battery baseline is infeasible under surplus_handling='must_absorb' because PV exceeds demand."
            )
        imported_power = max(0.0, demand - pv_limit)
        grid_import_series.append(imported_power)

    total_import_cost = sum(
        imported_power * price * timestep_hours
        for imported_power, price in zip(grid_import_series, input_dict["electricity_price"])
    )

    no_battery_peak_demand_cost = 0.0
    if peak_shaving_cost_factor > 0 and grid_import_series:
        if peak_shaving_frequency == "yearly":
            no_battery_peak_demand_cost = peak_shaving_cost_factor * max(grid_import_series)
        else:
            timestamps = pd.to_datetime(input_dict["timestamps"])
            df_temp = pd.DataFrame(
                {
                    "timestamp": timestamps,
                    "grid_import": grid_import_series,
                }
            )
            monthly_peaks = df_temp.groupby(df_temp["timestamp"].dt.to_period("M"))["grid_import"].max()
            no_battery_peak_demand_cost = peak_shaving_cost_factor * float(monthly_peaks.sum())

    return {
        "no_battery_import_cost": total_import_cost,
        "no_battery_peak_demand_cost": no_battery_peak_demand_cost,
        "no_battery_total_cost": total_import_cost + no_battery_peak_demand_cost,
    }

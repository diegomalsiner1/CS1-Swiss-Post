import numpy as np
from ortools.linear_solver import pywraplp
import pandas as pd
from tqdm import tqdm

## Will be used to define the optimization variables, constraints, and the objective function


def setup(input_dict, debug_infeasibility=False, fixed_battery_capacity_kwh=None):
    """
    Method for setting up the optimization problem

    Input: 
    Dict input_dict

    Output:
    ortools.linear_solver.pywraplp.solver optimization model
    """
    # Creating an optimization model/solver
    print(f"Setting up optimization problem")
    model = pywraplp.Solver.CreateSolver("CBC") # milp solver
    if model is None:
        raise RuntimeError("Could not create CBC solver.")

    model = pywraplp.Solver.CreateSolver("CBC")
    if model is None:
        raise RuntimeError(
            "Could not create CBC solver. Check OR-Tools installation and CBC backend availability."
        )

    ## Defining the required constants
    ## to be added
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
    optimization_mode = input_dict["parameters"].get("optimization_mode", "milp").lower()
    if optimization_mode not in {"milp", "lp"}:
        raise ValueError("optimization_mode must be 'milp' or 'lp'")
    use_binaries = optimization_mode == "milp"

    CRF = (((1+interest_rate)**lifetime) * interest_rate) / ((1+interest_rate)**lifetime - 1)

    ## Defining the timeseries of PV production, energy demand, and electricity prices


    ## Defining the required Variables
    Time_dependent_variables = {}
    # number of timesteps is driven by demand series length
    timesteps = range(len(input_dict["total_demand"]))

    inf = model.infinity()
    OPEX = model.NumVar(-inf, inf, f"Operational_Cost")
    Total_Cost = model.NumVar(-inf, inf, "Total_Cost")
    Battery_capacity = model.NumVar(0, Battery_capacity_upper_bound, "Battery_Capacity")
    slacks = []
    print("Initializing time dependent variables")
    battery_level_vars = []
    battery_in_flow_vars = []
    battery_out_flow_vars = []
    grid_flow_vars = []
    pv_out_flow_vars = []

    for t in tqdm(timesteps):
        # Integer Variables
        Time_dependent_variables[("Battery_in_flow",t)] = model.NumVar(0, Battery_max_inflow, f"Powerflow_Batter_in_{t}")
        Time_dependent_variables[("Battery_out_flow",t)] = model.NumVar(0, Battery_max_outflow, f"Powerflow_Batter_out_{t}")
        Time_dependent_variables[("Grid_flow",t)] = model.NumVar(0, inf, f"Powerflow_Grid_{t}")
        Time_dependent_variables[("Battery_level",t)] = model.NumVar(0, Battery_capacity_upper_bound, f"Battery_Level_{t}")
        Time_dependent_variables[("PV_out_flow",t)] = model.NumVar(0, PV_max_capacity, f"PV_Powerflow_out_{t}")

        battery_in_flow_vars.append(Time_dependent_variables[("Battery_in_flow", t)])
        battery_out_flow_vars.append(Time_dependent_variables[("Battery_out_flow", t)])
        battery_level_vars.append(Time_dependent_variables[("Battery_level", t)])
        grid_flow_vars.append(Time_dependent_variables[("Grid_flow", t)])
        pv_out_flow_vars.append(Time_dependent_variables[("PV_out_flow", t)])

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

        # PV Power Output
        pv_limit_t = input_dict["PV_capacity_factor"][t] * PV_max_capacity
        if debug_infeasibility:
            s_pv = model.NumVar(0, inf, f"slack_pv_limit_{t}")
            model.Add(Time_dependent_variables[("PV_out_flow", t)] <= pv_limit_t + s_pv)
            slacks.append(s_pv)
        else:
            model.Add(Time_dependent_variables[("PV_out_flow", t)] <= pv_limit_t)

        # MILP mode enforces explicit mutual exclusivity via binary on/off variables.
        if use_binaries:
            model.Add(Time_dependent_variables[("Binary_battery_in_flow",t)] + Time_dependent_variables[("Binary_battery_out_flow",t)] <= 1)
            model.Add(Time_dependent_variables[("Battery_in_flow",t)] <= Time_dependent_variables[("Binary_battery_in_flow",t)] * Battery_max_inflow)
            model.Add(Time_dependent_variables[("Battery_out_flow",t)] <= Time_dependent_variables[("Binary_battery_out_flow",t)] * Battery_max_outflow)
        model.Add(Time_dependent_variables[("Battery_level", t)] <= Battery_capacity)
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



    if fixed_battery_capacity_kwh is not None:
        if fixed_battery_capacity_kwh < 0 or fixed_battery_capacity_kwh > Battery_capacity_upper_bound:
            raise ValueError("fixed_battery_capacity_kwh must be within [0, Battery_max_capacity].")
        model.Add(Battery_capacity == fixed_battery_capacity_kwh)

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

    timestep_hours = 0.25
    import_cost_expr = sum(
        Time_dependent_variables[("Grid_flow", t)] * input_dict["electricity_price"][t] * timestep_hours
        for t in timesteps
    )
    annualized_battery_cost_expr = CRF * battery_invest_cost * Battery_capacity

    # Objective function
    if debug_infeasibility:
        model.Minimize(sum(slacks))
    else:
        opex_expr = cost_operation_and_maintenance + import_cost_expr
        model.Add(OPEX == opex_expr)
        model.Add(Total_Cost == annualized_battery_cost_expr + OPEX)
        model.Minimize(Total_Cost)

    solution_handles = {
        "battery_capacity": Battery_capacity,
        "opex": OPEX,
        "total_cost": Total_Cost,
        "annualized_battery_cost_expr": annualized_battery_cost_expr,
        "import_cost_expr": import_cost_expr,
        "fixed_om_cost": cost_operation_and_maintenance,
        "battery_level_vars": battery_level_vars,
        "battery_in_flow_vars": battery_in_flow_vars,
        "battery_out_flow_vars": battery_out_flow_vars,
        "grid_flow_vars": grid_flow_vars,
        "pv_out_flow_vars": pv_out_flow_vars,
    }

    return model, slacks, solution_handles


def optimize_model(model, slacks=None, top_n=20):
    """
    Method for optimizing the a constrainted model
    Input:
    ortools.linear_solver.pywraplp.solver unsolved optimization model

    Output:
    ortools.linear_solver.pywraplp.solver solved optimization model
    """
    print("Running optimization")
    model.EnableOutput()  # Enable solver output for debugging
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
    return {
        "battery_capacity_kwh": solution_handles["battery_capacity"].solution_value(),
        "objective_total_cost": model.Objective().Value(),
        "opex": solution_handles["opex"].solution_value(),
        "import_cost": solution_handles["import_cost_expr"].solution_value(),
        "fixed_om_cost": solution_handles["fixed_om_cost"],
        "annualized_battery_cost": solution_handles["annualized_battery_cost_expr"].solution_value(),
    }


def compute_no_battery_baseline(input_dict):
    """Compute the baseline import cost with no battery (all PV used to satisfy demand)."""
    timestep_hours = 0.25
    if len(input_dict.get("total_demand", [])) != len(input_dict.get("PV_capacity_factor", [])):
        raise ValueError("Input timeseries lengths must match for baseline computation.")

    total_import_cost = 0.0
    for t, demand in enumerate(input_dict["total_demand"]):
        pv_limit = input_dict["PV_capacity_factor"][t] * input_dict["parameters"]["PV_max_capacity"]
        imported_energy = max(0.0, demand - pv_limit)
        price = input_dict["electricity_price"][t]
        total_import_cost += imported_energy * price * timestep_hours

    no_battery_opex = input_dict["parameters"].get("operation_and_maintenance", 0.0)
    return {
        "no_battery_import_cost": total_import_cost,
        "no_battery_total_cost": no_battery_opex + total_import_cost,
    }

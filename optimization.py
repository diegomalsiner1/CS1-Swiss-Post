import numpy as np
from ortools.linear_solver import pywraplp
import pandas as pd
from tqdm import tqdm

## Will be used to define the optimization variables, constraints, and the objective function


def setup(input_dict, debug_infeasibility=False):
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
    Battery_max_capacity = input_dict["parameters"]["Battery_max_capacity"]
    eta_charge = input_dict["parameters"]["Battery_eta_charge"]
    eta_discharge = input_dict["parameters"]["Battery_eta_discharge"]
    eta_self_discharge = input_dict["parameters"]["Battery_eta_self_discharge"]
    battery_invest_cost = input_dict["parameters"]["Battery_invest_cost"]
    cost_operation_and_maintenance = input_dict["parameters"]["operation_and_maintenance"]

    interest_rate = input_dict["parameters"]["interest_rate"]
    lifetime = input_dict["parameters"]["lifetime"]
    battery_degrading = input_dict["parameters"]["battery_degrading"]

    CRF = (((1+interest_rate)**lifetime) * interest_rate) / ((1+interest_rate)**lifetime - 1)

    ## Defining the timeseries of PV production, energy demand, and electricity prices


    ## Defining the required Variables
    Time_dependent_variables = {}
    # number of timesteps is driven by demand series length
    timesteps = range(len(input_dict["total_demand"]))

    inf = model.infinity()
    OPEX = model.NumVar(-inf, inf, f"Operational_Cost")
    slacks = []
    print("Initializing time dependent variables")
    for t in tqdm(timesteps):
        # Integer Variables
        Time_dependent_variables[("Battery_in_flow",t)] = model.NumVar(0, Battery_max_inflow, f"Powerflow_Batter_in_{t}")
        Time_dependent_variables[("Battery_out_flow",t)] = model.NumVar(0, Battery_max_outflow, f"Powerflow_Batter_out_{t}")
        Time_dependent_variables[("Grid_flow",t)] = model.NumVar(0, inf, f"Powerflow_Grid_{t}")
        Time_dependent_variables[("Battery_level",t)] = model.NumVar(0, Battery_max_capacity, f"Battery_Level_{t}")
        Time_dependent_variables[("PV_out_flow",t)] = model.NumVar(0, PV_max_capacity, f"PV_Powerflow_out_{t}")

        # Binary Variables
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

        # Equations for battery inflow and outflow
        model.Add(Time_dependent_variables[("Binary_battery_in_flow",t)] + Time_dependent_variables[("Binary_battery_out_flow",t)] <= 1)
        model.Add(Time_dependent_variables[("Battery_in_flow",t)] <= Time_dependent_variables[("Binary_battery_in_flow",t)] * Battery_max_inflow)
        model.Add(Time_dependent_variables[("Battery_out_flow",t)] <= Time_dependent_variables[("Binary_battery_out_flow",t)] * Battery_max_outflow)
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
        model.Add(Time_dependent_variables[("Battery_level", 0)] - 0.5 * Battery_max_capacity <= s_soc_start)
        model.Add(0.5 * Battery_max_capacity - Time_dependent_variables[("Battery_level", 0)] <= s_soc_start)
        model.Add(Time_dependent_variables[("Battery_level", timesteps[-1])] - 0.5 * Battery_max_capacity <= s_soc_end)
        model.Add(0.5 * Battery_max_capacity - Time_dependent_variables[("Battery_level", timesteps[-1])] <= s_soc_end)
        slacks.extend([s_soc_start, s_soc_end])
    else:
        model.Add(Time_dependent_variables[("Battery_level", 0)] == 0.5 * Battery_max_capacity)
        model.Add(Time_dependent_variables[("Battery_level", timesteps[-1])] == 0.5 * Battery_max_capacity)

    import_cost_expr = sum(
        Time_dependent_variables[("Grid_flow", t)] * input_dict["electricity_price"][t]
        for t in timesteps
    )

    # Objective function
    if debug_infeasibility:
        model.Minimize(sum(slacks))
    else:
        model.Add(OPEX == CRF * battery_invest_cost * Battery_max_capacity + battery_degrading + cost_operation_and_maintenance + import_cost_expr)
        model.Minimize(OPEX)

    return model, slacks


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

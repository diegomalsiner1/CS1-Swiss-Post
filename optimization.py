import numpy as np
from ortools.linear_solver import pywraplp
import pandas as pd
from tqdm import tqdm

## Will be used to define the optimization variables, constraints, and the objective function


def setup(input_dict): 
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
    timesteps = range(input_dict.size()) ## to be changed
    

    OPEX = model.NewIntVar(-np.INF, np.INF,f"Operational_Cost")
    Import_Cost = model.NewIntVar(-np.INF, np.INF,f"Import_Cost")
    print("Initializing time dependent variables")
    for t in tqdm(timesteps):
        # Integer Variables
        Time_dependent_variables[("Battery_in_flow",t)] = model.NewIntVar(0, Battery_max_inflow, f"Powerflow_Batter_in_{t}")
        Time_dependent_variables[("Battery_out_flow",t)] = model.NewIntVar(0, Battery_max_outflow, f"Powerflow_Batter_out_{t}")
        Time_dependent_variables[("Grid_flow",t)] = model.NewIntVar(0, np.INF, f"Powerflow_Grid_{t}")
        Time_dependent_variables[("Battery_level",t)] = model.NewIntVar(0, Battery_max_capacity, f"Battery_Level_{t}")
        Time_dependent_variables[("PV_out_flow",t)] = model.NewIntVar(0, PV_max_capacity, f"PV_Powerflow_out_{t}")

        # Binary Variables
        Time_dependent_variables[("Binary_battery_in_flow",t)] = model.NewBoolVar(f"Binary_battery_in_flow")
        Time_dependent_variables[("Binary_battery_out_flow",t)] = model.NewBoolVar(f"Binary_battery_out_flow")

    ## Constraint functions
    # time dependent constraints
    print("Intitializing time dependent constraint functions")
    for t in tqdm(timesteps):
        # Power Flow Balance
        model.Add(0 == Time_dependent_variables[("Grid_flow",t)] + Time_dependent_variables[("PV_out_flow",t)] + Time_dependent_variables[("Battery_out_flow",t)] - Time_dependent_variables[("Battery_in_flow",t)] - input_dict["total_demand",t])
        # PV Power Output
        model.Add(Time_dependent_variables[("PV_out_flow",t)] == input_dict["PV_capacity_factor",t] * PV_max_capacity)

        # Equations for battery inflow and outflow
        model.Add(Time_dependent_variables[("Binary_battery_in_flow",t)] + Time_dependent_variables[("Binary_battery_out_flow",t)] <= 1)
        model.Add(Time_dependent_variables[("Battery_in_flow",t)] <= Time_dependent_variables[("Binary_battery_in_flow",t)] * Battery_max_inflow)
        model.Add(Time_dependent_variables[("Battery_out_flow",t)] <= Time_dependent_variables[("Binary_battery_out_flow",t)] * Battery_max_outflow)
        if t != timesteps[-1]:
            # Equation to calculate the battery level in the next time step t+1 based on the current time step t
            model.Add(Time_dependent_variables[("Battery_level",t+1)] == Time_dependent_variables[("Battery_level",t)] * (1-eta_self_discharge) + 0.25 * eta_charge * Time_dependent_variables[("Battery_in_flow",t)] - 0.25 * 1/eta_discharge * Time_dependent_variables[("Battery_out_flow",t)])



    # starting conditions
    model.Add(Time_dependent_variables[("Battery_level",0)] == 0.5 * Battery_max_capacity)
    model.Add(Time_dependent_variables[("Battery_level",timesteps[-1])] == 0.5 * Battery_max_capacity)
    Import_Cost = sum(
    Time_dependent_variables[("Grid_flow", t)] * input_dict[("electricity_price", t)]
    for t in timesteps
)

    # Objective function
    model.Add(OPEX == CRF * battery_invest_cost * Battery_max_capacity + battery_degrading + cost_operation_and_maintenance + Import_Cost) 
    model.Minimize(OPEX)
  
    return model


def optimize_model(model):
    """
    Method for optimizing the a constrainted model
    Input:
    ortools.linear_solver.pywraplp.solver unsolved optimization model

    Output:
    ortools.linear_solver.pywraplp.solver solved optimization model
    """

    status = model.Solve()

    if status == model.OPTIMAL:
        print('Optimal solution found')
   
    else:
        raise ValueError(f'Did not find optimal solution: {status}')
    return model

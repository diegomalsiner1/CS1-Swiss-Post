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

    ## Defining the timeseries of PV production, energy demand, and electricity prices


    ## Defining the required Variables
    Time_dependent_variables = {}
    timesteps = range(input_dict.size()) ## to be changed
    

    NPV = model.NewIntVar(-np.INF, np.INF,f"Net_Present_Value")
    print("Initializing time dependent variables")
    for t in tqdm(timesteps):
        # Integer Variables
        Time_dependent_variables[("Battery_in_flow",t)] = model.NewIntVar(0, Battery_max_inflow, f"Powerflow_Batter_in_{t}")
        Time_dependent_variables[("Battery_out_flow",t)] = model.NewIntVar(0, Battery_max_outflow, f"Powerflow_Batter_out_{t}")
        Time_dependent_variables[("Grid_flow",t)] = model.NewIntVar(0, np.INF, f"Powerflow_Grid_{t}")
        Time_dependent_variables[("Battery_level",t)] = model.NewIntVar(0, Battery_max_capacity, f"Battery_Level_{t}")

        # Binary Variables
        Time_dependent_variables[("Binary_battery_in_flow",t)] = model.NewBoolVar(f"Binary_battery_in_flow")
        Time_dependent_variables[("Binary_battery_out_flow",t)] = model.NewBoolVar(f"Binary_battery_out_flow")

    ## Constraint functions
    # time dependent constraints
    print("Intitializing time dependent constraint functions")
    for t in tqdm(timesteps):
        # Power Flow Balance
        ## to be added

        # Equations for battery inflow and outflow
        model.Add(Time_dependent_variables[("Binary_battery_in_flow",t)] + Time_dependent_variables[("Binary_battery_out_flow",t)] <= 1)
        model.Add(Time_dependent_variables[("Battery_in_flow",t)] <= Time_dependent_variables[("Binary_battery_in_flow",t)] * Battery_max_inflow)
        model.Add(Time_dependent_variables[("Battery_out_flow",t)] <= Time_dependent_variables[("Binary_battery_out_flow",t)] * Battery_max_outflow)
        if t != timesteps[-1]:
            # Equation to calculate the battery level in the next time step t+1 based on the current time step t
            model.Add(Time_dependent_variables[("Battery_level",t+1)] == Time_dependent_variables[("Battery_level",t)] * (1-eta_self_discharge) + 0.25 * eta_charge * Time_dependent_variables[("Battery_in_flow",t)] - 0.25 * 1/eta_discharge * Time_dependent_variables[("Battery_out_flow",t)])

    # time independent constraints
    model.Add(Time_dependent_variables[("Battery_level",0)] == 0.5 * Battery_max_capacity)
    model.Add(Time_dependent_variables[("Battery_level",timesteps[-1])] == 0.5 * Battery_max_capacity)


    # Objective function
    model.Add(NPV == 4) ## to be changed
    model.Maximize(NPV)

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
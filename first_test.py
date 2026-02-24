
import numpy as np
from ortools.linear_solver import pywraplp
import pandas as pd

T = range(24)           # for testing only 24 hours 
INF = 1000000000          ## for boundaries


## some dummy values for testing
P_base_load = [
    18, 24, 29, 22, 28, 35, 70, 80, 110, 150, 200, 220, 225,
    210, 200, 180, 150, 130, 80, 55, 32, 19, 21, 29
]
#P_base_load = pd.from_csv("path.csv")


P_EV_load = [
    190, 185, 174, 170, 124, 155, 147, 111, 119, 121, 99, 72,
    80, 95, 110, 120, 130, 125, 115, 105, 94, 120, 140, 170
]
#P_EV_load = pd.from_csv("path.csv")

PV_capacity_factor = [
    0.0, 0.0, 0.0, 0.0, 0.0, 0.05, 0.1, 0.25, 0.4, 0.5, 0.6,
    0.65, 0.65, 0.52, 0.39, 0.25, 0.12, 0.06, 0.02, 0.0, 0.0, 0.0, 0.0, 0.0
]

k_el = [
    11.6, 12.6, 13.6, 14.1, 13.6, 11.6, 10.6, 10.1,  9.8,  9.6, 9.5, 9.6, 10.6, 12.1,
    12.6, 13.1, 12.8, 12.0,  11.6, 11.8, 12.3, 12.6, 13.1, 13.6
]


solver = pywraplp.Solver.CreateSolver("CBC") # milp solver



P_pv = list(solver.NumVar(lb=0, ub=INF, name=f'P_PV({t})') for t in T)
P_total_demand = list(solver.NumVar(lb=0, ub=INF, name=f'P_total_demand({t})') for t in T)
P_grid = list(solver.NumVar(lb=0, ub=INF, name=f'P_grid({t})') for t in T)


P_battery_in = list(solver.NumVar(lb=0, ub=INF, name=f'P_battery_in({t})') for t in T)
P_battery_out = list(solver.NumVar(lb=0, ub=INF, name=f'P_battery_out({t})') for t in T)
E_battery = list(solver.NumVar(lb=0, ub=INF, name=f'E_battery({t})') for t in T)

x_battery_in = list(solver.BoolVar(f"x_bat_in[{t}]") for t in T)
x_battery_out =list(solver.BoolVar(f"x_bat_out[{t}]") for t in T)

OPEX = list(solver.NumVar(lb=0, ub=INF, name=f'OPEX({t})') for t in T)

P_PV_Peak = 300
E_battery_max = 1000
eta_battery_in = 0.9
eta_battery_out = 0.9
eta_battery_self_discharge = 0.01
P_battery_max_flow = 10

for t in T:
    
    #Constraints
    solver.Add(P_total_demand[t] == P_base_load[t] + P_EV_load[t])

    solver.Add(P_total_demand[t] == P_pv[t] + P_battery_out[t]  - P_battery_in[t] + P_grid[t])
    if t != 23:
        solver.Add(E_battery[t+1] == E_battery[t] * (1-eta_battery_self_discharge) + P_battery_in[t] * eta_battery_in + P_battery_out[t] / eta_battery_out)
    solver.Add(P_pv[t] == P_PV_Peak * PV_capacity_factor[t])
    solver.Add(OPEX[t] == P_grid[t] * k_el[t])

    solver.Add(P_battery_in[t] <= P_battery_max_flow * x_battery_in[t])
    solver.Add(P_battery_out[t] <= P_battery_max_flow * x_battery_out[t])
    solver.Add(x_battery_in[t] + x_battery_out[t] <= 1)
    solver.Add(E_battery[t] <= E_battery_max)
    # if electricity can only be bought
    solver.Add(P_grid[t] >= 0)

solver.Add(E_battery[0] == 0.25 * E_battery_max) # starting condition
solver.Add(E_battery[23] == 0.25 * E_battery_max) # starting condition
solver.Minimize(sum(OPEX))

status = solver.Solve()

if status == solver.OPTIMAL:
    print('Problem is solvable')
    print(f'The objective function results in {np.round(solver.Objective().Value(), 2)}')
    print('The variables have the following values:')
    for var in solver.variables():
        print(f'\t{var.name()}= {np.round(var.solution_value(), 2)}')
    
else:
    raise ValueError(f'Did not find optimal solution: {status}')
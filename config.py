## Will be used to define the input parameters

use_case = "Peak_Shaving" ## can be switched between "Peak_Shaving"
interest_rate = 0.06 # [-]
lifetime = 20 # [years]
year =  2025 # 

load_existing_input_dict = True # [True or False]

PV_max_capacity = 10000 # [kW]
Battery_max_inflow = 10000 # [kW]
Battery_max_outflow = 10000 # [kW]
Battery_max_capacity = 1000000 # [kWh]
eta_charge = 0.9 # [-]
eta_discharge = 0.95 # [-]
eta_self_discharge = 0.01 # [-]
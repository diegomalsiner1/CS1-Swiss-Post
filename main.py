import pandas as pd
import numpy as np
import optimization as opt
import data_preprocessing as dpp
import results_processing as rp
import config

## Will be used to run the tool


# ==========================================================
# Data Preprocessing
# ==========================================================
dfs = []
# Extract trafo data 
trafo_sheets = dpp.select_sheets("Select sheets with trafo data:")
trafo_df = dpp.load_grid_exchange(trafo_sheets)
dfs.append(trafo_df)

# PV data (or just trafo2?)
pv_sheets = dpp.select_sheets("Select the PV sheet")
pv_df = dpp.load_trafo(pv_sheets[0])
pv_df = pv_df.rename(columns={"power_kW": "PV_kW"})
dfs.append(pv_df)

# ev data
lkw = dpp.generate_lkw_profile(year=2025)
zustellung = dpp.generate_zustellung_profile(year=2025)
ev_total = lkw.merge(zustellung, on="timestamp", how="outer")
dfs.append(ev_total)


# Merge all on 'timestamp' and export to csv
merged_df = reduce(lambda left, right: pd.merge(left, right, on='timestamp', how="outer"), dfs)
merged_df = merged_df.sort_values('timestamp').reset_index(drop=True)
merged_df.to_csv("03-PROCESSED-DATA/data_processed.csv", index=False)

# Build optimization input dictionary
merged_df["lkw_kW"] = merged_df.get("lkw_kW", 0).fillna(0)
merged_df["zustellung_kW"] = merged_df.get("zustellung_kW", 0).fillna(0)
merged_df["grid_exchange_kW"] = merged_df.get("grid_exchange_kW", 0).fillna(0)
merged_df["PV_kW"] = merged_df.get("PV_kW", 0).fillna(0)

# Total demand in kW (adjust with your local definition)
merged_df["total_demand"] = merged_df["lkw_kW"] + merged_df["zustellung_kW"] + merged_df["grid_exchange_kW"]

# PV capacity factor in [0,1]
PV_max = config.PV_max_capacity
merged_df["PV_capacity_factor"] = (merged_df["PV_kW"] / PV_max).clip(lower=0, upper=1)

# konstantes Preisprofil (30 Rappen/kWh) – ersetzen Sie es ggf. durch echte Kurve
price_chf_per_kwh = 0.30
merged_df["electricity_price"] = price_chf_per_kwh

input_dict = {
    "parameters": {
        "PV_max_capacity": config.PV_max_capacity,
        "Battery_max_inflow": config.Battery_max_inflow,
        "Battery_max_outflow": config.Battery_max_outflow,
        "Battery_max_capacity": config.Battery_max_capacity,
        "Battery_eta_charge": config.eta_charge,
        "Battery_eta_discharge": config.eta_discharge,
        "Battery_eta_self_discharge": config.eta_self_discharge,
        "Battery_invest_cost": 450,  # CHF/kWh, adaptiere nach config
        "operation_and_maintenance": 10000,  # CHF/a, placeholder
        "interest_rate": config.interest_rate,
        "lifetime": config.lifetime,
        "battery_degrading": 0  # placeholder
    },
    "total_demand": merged_df["total_demand"].tolist(),
    "PV_capacity_factor": merged_df["PV_capacity_factor"].tolist(),
    "electricity_price": merged_df["electricity_price"].tolist()
}

# Run optimization
model = opt.setup(input_dict)
model = opt.optimize_model(model)

print("Optimization finished")
print("Objective OPEX", model.Objective().Value())

#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

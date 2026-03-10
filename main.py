import pandas as pd
import numpy as np
import optimization as opt
import data_preprocessing as dpp
import results_processing as rp
import config
from functools import reduce

## Will be used to run the tool


# ==========================================================
# Data Preprocessing
# ==========================================================
dfs = []
# Extract trafo data 
trafo_sheets = dpp.select_sheets("Select sheets with grid exchange data:")
trafo_df = dpp.load_grid_exchange(trafo_sheets)
trafo_df = trafo_df[["timestamp", "grid_exchange_kW"]]
dfs.append(trafo_df)

# PV data
pv_sheets = dpp.select_sheets("Select the PV sheet")
pv_df = dpp.load_grid_exchange(pv_sheets)
pv_df = pv_df[["timestamp", "grid_exchange_kW"]].rename(columns={"grid_exchange_kW": "PV_kW"})
dfs.append(pv_df)

# ev data
ev_df = dpp.generate_lkw_profile(year=2024)
dfs.append(ev_df)

# zustellung data
zustellung_df = dpp.generate_zustellung_profile(year=2024)
dfs.append(zustellung_df)

# feed in tarif



# Merge all on 'timestamp', create total and export to csv
merged_df = reduce(lambda left, right: pd.merge(left, right, on='timestamp', how="outer"), dfs)
merged_df["total_kW"] = merged_df.drop(columns="timestamp").sum(axis=1)
merged_df["feed_in_tariff_CHF_per_kWh"] = 0.12 # temp constant feed in
merged_df.to_csv("03-PROCESSED-DATA/data_processed.csv", index=False)

#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

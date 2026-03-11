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
merged_df.to_csv("03-PROCESSED-DATA/data_processed.csv", index=False)

#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

import pandas as pd
import numpy as np
from functools import reduce
import json
from pathlib import Path
import optimization as opt
import data_preprocessing as dpp
import results_processing as rp
import config

## Will be used to run the tool


INPUT_DICT_PATH = Path("03-PROCESSED-DATA/input_dict.json")
DEBUG_INFEASIBILITY = True


def apply_timestep_cap(input_dict: dict, max_timesteps) -> dict:
    """
    Truncate all time-series entries in input_dict to max_timesteps.
    """
    if max_timesteps is None:
        return input_dict

    if not isinstance(max_timesteps, int) or max_timesteps <= 0:
        raise ValueError("config.max_timesteps must be a positive integer or None")

    series_keys = ["total_demand", "PV_capacity_factor", "electricity_price"]
    available_lengths = [len(input_dict[k]) for k in series_keys if k in input_dict]
    if not available_lengths:
        return input_dict

    current_steps = min(available_lengths)
    target_steps = min(current_steps, max_timesteps)

    if target_steps == current_steps:
        return input_dict

    for key in series_keys:
        if key in input_dict:
            input_dict[key] = input_dict[key][:target_steps]

    print(f"Applied timestep cap: {current_steps} -> {target_steps}")
    return input_dict


def build_input_dict_from_raw_data() -> dict:
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

    # EV data
    lkw = dpp.generate_lkw_profile(year=2025)
    zustellung = dpp.generate_zustellung_profile(year=2025)
    ev_total = lkw.merge(zustellung, on="timestamp", how="outer")
    dfs.append(ev_total)

    # Merge all on 'timestamp' and export to csv
    merged_df = reduce(lambda left, right: pd.merge(left, right, on="timestamp", how="outer"), dfs)
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

    # Constant price profile (CHF/kWh)
    price_chf_per_kwh = 0.30
    merged_df["electricity_price"] = price_chf_per_kwh

    return {
        "parameters": {
            "PV_max_capacity": config.PV_max_capacity,
            "Battery_max_inflow": config.Battery_max_inflow,
            "Battery_max_outflow": config.Battery_max_outflow,
            "Battery_max_capacity": config.Battery_max_capacity,
            "Battery_eta_charge": config.eta_charge,
            "Battery_eta_discharge": config.eta_discharge,
            "Battery_eta_self_discharge": config.eta_self_discharge,
            "Battery_invest_cost": config.invest_cost,
            "operation_and_maintenance": config.operation_and_maintenance,
            "interest_rate": config.interest_rate,
            "lifetime": config.lifetime,
            "battery_degrading": config.battery_degrading
        },
        "total_demand": merged_df["total_demand"].tolist(),
        "PV_capacity_factor": merged_df["PV_capacity_factor"].tolist(),
        "electricity_price": merged_df["electricity_price"].tolist()
    }


def load_or_build_input_dict() -> dict:
    if config.load_existing_input_dict:
        if not INPUT_DICT_PATH.exists():
            raise FileNotFoundError(
                f"load_existing_input_dict=True, but no file found at '{INPUT_DICT_PATH}'. "
                "Run once with load_existing_input_dict=False to generate it."
            )
        with INPUT_DICT_PATH.open("r", encoding="utf-8") as f:
            print(f"Loading input dictionary from {INPUT_DICT_PATH}")
            input_dict = json.load(f)

        steps_before = len(input_dict.get("total_demand", []))
        input_dict = apply_timestep_cap(input_dict, config.max_timesteps)
        steps_after = len(input_dict.get("total_demand", []))
        if steps_after < steps_before:
            with INPUT_DICT_PATH.open("w", encoding="utf-8") as f:
                json.dump(input_dict, f)
            print(f"Updated {INPUT_DICT_PATH} with capped time series")
        return input_dict

    input_dict = build_input_dict_from_raw_data()
    input_dict = apply_timestep_cap(input_dict, config.max_timesteps)
    INPUT_DICT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with INPUT_DICT_PATH.open("w", encoding="utf-8") as f:
        json.dump(input_dict, f)
    print(f"Saved input dictionary to {INPUT_DICT_PATH}")
    return input_dict


input_dict = load_or_build_input_dict()

# Run optimization
model, slacks = opt.setup(input_dict, debug_infeasibility=DEBUG_INFEASIBILITY)
model = opt.optimize_model(model, slacks=slacks if DEBUG_INFEASIBILITY else None)

print("Optimization finished")
print("Objective OPEX", model.Objective().Value())
#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

import pandas as pd
import numpy as np
from functools import reduce
import json
from datetime import datetime
from pathlib import Path
import time
import optimization as opt
import data_preprocessing as dpp
import results_processing as rp
import config

## Will be used to run the tool


INPUT_DICT_PATH = Path("03-PROCESSED-DATA/input_dict.json")
RESULTS_ROOT = Path("02-MODEL-RESULTS")
DEBUG_INFEASIBILITY = False


def create_run_output_dir(input_dict: dict) -> Path:
    mode = input_dict.get("parameters", {}).get("optimization_mode", "unknown")
    n_steps = len(input_dict.get("total_demand", []))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = RESULTS_ROOT / f"{stamp}_{mode}_{n_steps}steps"
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def save_run_artifacts(run_dir: Path, input_dict: dict, solution_summary: dict, runtime_s: float) -> None:
    settings_snapshot = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "debug_infeasibility": DEBUG_INFEASIBILITY,
        "load_existing_input_dict": config.load_existing_input_dict,
        "max_timesteps": config.max_timesteps,
        "timesteps_used": len(input_dict.get("total_demand", [])),
        "parameters": input_dict.get("parameters", {}),
    }

    results_summary = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "runtime_seconds": runtime_s,
        "objective_total_cost": float(solution_summary["objective_total_cost"]),
        "battery_capacity_kwh": float(solution_summary["battery_capacity_kwh"]),
        "opex": float(solution_summary["opex"]),
        "import_cost": float(solution_summary["import_cost"]),
        "fixed_om_cost": float(solution_summary["fixed_om_cost"]),
        "annualized_battery_cost": float(solution_summary["annualized_battery_cost"]),
        "peak_demand_cost": float(solution_summary.get("peak_demand_cost", 0.0)),
        "no_battery_import_cost": float(solution_summary.get("no_battery_import_cost", 0.0)),
        "no_battery_total_cost": float(solution_summary.get("no_battery_total_cost", 0.0)),
        "npv": float(solution_summary.get("npv", np.nan)),
        "irr": float(solution_summary.get("irr", np.nan)),
        "payback_years": float(solution_summary.get("payback_years", np.nan)),
    }

    (run_dir / "settings_snapshot.json").write_text(
        json.dumps(settings_snapshot, indent=2),
        encoding="utf-8",
    )
    (run_dir / "results_summary.json").write_text(
        json.dumps(results_summary, indent=2),
        encoding="utf-8",
    )


def apply_timestep_cap(input_dict: dict, max_timesteps) -> dict:
    """
    Truncate all time-series entries in input_dict to max_timesteps.
    """
    if max_timesteps is None:
        return input_dict

    if not isinstance(max_timesteps, int) or max_timesteps <= 0:
        raise ValueError("config.max_timesteps must be a positive integer or None")

    series_keys = ["timestamps", "total_demand", "PV_capacity_factor", "electricity_price"]
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
            "battery_degrading": config.battery_degrading,
            "optimization_mode": config.optimization_mode,
            "peak_shaving_cost_factor": config.peak_shaving_cost_factor,
            "peak_shaving_granularity": config.peak_shaving_granularity,
        },
        "timestamps": merged_df["timestamp"].astype(str).tolist(),
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
run_dir = create_run_output_dir(input_dict)
run_start = time.perf_counter()

# Run optimization
model, slacks, solution_handles = opt.setup(input_dict, debug_infeasibility=DEBUG_INFEASIBILITY)
model = opt.optimize_model(model, slacks=slacks if DEBUG_INFEASIBILITY else None, debug_infeasibility=DEBUG_INFEASIBILITY)
solution_summary = opt.summarize_solution(model, solution_handles)
run_seconds = time.perf_counter() - run_start
baseline_summary = opt.compute_no_battery_baseline(input_dict)
solution_summary.update(baseline_summary)
save_run_artifacts(run_dir, input_dict, solution_summary, run_seconds)

battery_soc = [v.solution_value() for v in solution_handles["battery_level_vars"]]
pv_flow = [v.solution_value() for v in solution_handles["pv_out_flow_vars"]]
grid_flow = [v.solution_value() for v in solution_handles["grid_flow_vars"]]
total_load = input_dict.get("total_demand", [])
timestamps = input_dict.get("timestamps")
rp.export_results(
    run_dir,
    solution_summary,
    battery_soc,
    timestamps=timestamps,
    input_dict=input_dict,
    pv_flow=pv_flow,
    grid_flow=grid_flow,
    total_load=total_load,
)

print("Optimization finished")
print("Objective Total Cost", solution_summary["objective_total_cost"])
print("Battery Capacity [kWh]", solution_summary["battery_capacity_kwh"])
print("OPEX", solution_summary["opex"])
print("Import Cost", solution_summary["import_cost"])
print("Fixed O&M Cost", solution_summary["fixed_om_cost"])
print("Annualized Battery Cost", solution_summary["annualized_battery_cost"])
print("Runtime [s]", run_seconds)
print("Saved run artifacts to", run_dir)
#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

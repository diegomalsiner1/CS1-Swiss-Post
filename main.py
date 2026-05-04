import pandas as pd
import numpy as np
from functools import reduce
import json
import importlib
from datetime import datetime
from pathlib import Path
import config
import time
import optimization as opt
import results_processing as rp
import report_generation as report_gen

## Will be used to run the tool
INPUT_DICT_PATH = Path("03-PROCESSED-DATA/input_dict.json")
RESULTS_ROOT = Path("02-MODEL-RESULTS")
DEBUG_INFEASIBILITY = False


def get_runtime_parameters() -> dict:
    return {
            # --- Standard Parameters ---
            "PV_max_capacity": config.PV_max_capacity,
            "Battery_max_inflow": config.Battery_max_inflow,
            "Battery_max_outflow": config.Battery_max_outflow,
            "Battery_max_capacity": config.Battery_max_capacity,
            "Battery_eta_charge": config.eta_charge,
            "Battery_eta_discharge": config.eta_discharge,
            "Battery_eta_self_discharge": config.eta_self_discharge,
            "operation_and_maintenance": config.operation_and_maintenance,
            "interest_rate": config.interest_rate,
            "lifetime": config.lifetime,
            "battery_degrading": config.battery_degrading,
            "optimization_mode": config.optimization_mode,
            "peak_shaving_cost_factor": config.peak_shaving_cost_factor,
            "peak_shaving_frequency": config.peak_shaving_frequency,

            # Advanced Parameters ---
            # Using getattr to safely handle new Excel columns or config variables
            "Battery_invest_cost": getattr(config, "invest_cost", 0.0),
            "Battery_energy_invest_cost": getattr(config, "invest_cost_energy", getattr(config, "invest_cost", 0.0)),
            "Battery_power_invest_cost": getattr(config, "invest_cost_power", 0.0),
            "battery_max_c_rate": getattr(config, "battery_max_c_rate", None),
            "battery_min_soc_fraction": getattr(config, "battery_min_soc_fraction", 0.0),
            "battery_cycle_life": getattr(config, "battery_cycle_life", None),
            "battery_calendar_life_years": getattr(config, "battery_calendar_life_years", None),
            "surplus_handling": getattr(config, "surplus_handling", "curtail"),
            "battery_replacement_cost_fraction": getattr(config, "battery_replacement_cost_fraction", 0.0),
            # Parameters for sensitivity and report
            "run_battery_size_sensitivity": getattr(config, "run_battery_size_sensitivity", False),
            "battery_sensitivity_sizes_kwh": getattr(config, "battery_sensitivity_sizes_kwh", []),
            "generate_pdf_report": getattr(config, "generate_pdf_report", False),
        }

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
        "battery_power_capacity_kw": float(solution_summary.get("battery_power_capacity_kw", 0.0)),
        "opex": float(solution_summary["opex"]),
        "import_cost": float(solution_summary["import_cost"]),
        "fixed_om_cost": float(solution_summary["fixed_om_cost"]),
        "annualized_battery_cost": float(solution_summary["annualized_battery_cost"]),
        "peak_demand_cost": float(solution_summary.get("peak_demand_cost", 0.0)),
        "no_battery_import_cost": float(solution_summary.get("no_battery_import_cost", 0.0)),
        "no_battery_peak_demand_cost": float(solution_summary.get("no_battery_peak_demand_cost", 0.0)),
        "no_battery_total_cost": float(solution_summary.get("no_battery_total_cost", 0.0)),
        "curtailed_energy_kwh": float(solution_summary.get("curtailed_energy_kwh", 0.0)),
        "replacement_cost": float(solution_summary.get("replacement_cost", 0.0)),
        "replacement_year": solution_summary.get("replacement_year"),
        "npv": float(solution_summary.get("npv", np.nan)),
        "irr": float(solution_summary.get("irr", np.nan)),
        "payback_years": float(solution_summary.get("payback_years", np.nan)),
        "discounted_payback_years": float(solution_summary.get("discounted_payback_years", np.nan)),
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
    import data_preprocessing as dpp

    # ==========================================================
    # Data Preprocessing
    # ==========================================================
    dfs = []

    # Extract trafo data
    trafo_sheets = dpp.select_sheets("Select sheets with trafo data:")
    trafo_df = dpp.load_grid_exchange(trafo_sheets)
    year = trafo_df["timestamp"].dt.year.iloc[0]
    dfs.append(trafo_df)

    # PV data (or just trafo2?)
    pv_sheets = dpp.select_sheets("Select the PV sheet")
    pv_df = dpp.load_trafo(pv_sheets[0])
    pv_df = pv_df.rename(columns={"power_kW": "PV_kW"})
    dfs.append(pv_df)

    # EV data
    ev_ChargingSheet = dpp.select_sheets("Select the EV charging sheet")
    dist_sheet = dpp.select_sheets("Select the sheet with the distribution profile")
    lkw = dpp.generate_lkw_profile(year=year, sheet_name=ev_ChargingSheet[0])
    zustellung = dpp.generate_zustellung_profile(year=year, sheet_name=dist_sheet[0])
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

    # Dynamic price profile from ENTSO-E day-ahead spot prices (CHF/kWh)
    price_df = dpp.load_price_curve(year=year)
    merged_df = merged_df.merge(price_df, on="timestamp", how="left")
    merged_df["electricity_price"] = merged_df["electricity_price"].ffill().bfill()

    return {
        "parameters": get_runtime_parameters(),
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
        input_dict["parameters"] = get_runtime_parameters()

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


def run_battery_size_sensitivity(input_dict: dict, battery_sizes_kwh) -> pd.DataFrame:
    rows = []
    baseline_summary = opt.compute_no_battery_baseline(input_dict)
    invest_cost_per_kwh = float(
        input_dict["parameters"].get(
            "Battery_energy_invest_cost",
            input_dict["parameters"].get("Battery_invest_cost", 0.0),
        )
    )
    invest_cost_per_kw = float(input_dict["parameters"].get("Battery_power_invest_cost", 0.0))
    interest_rate = float(input_dict["parameters"].get("interest_rate", 0.0))
    lifetime = int(input_dict["parameters"].get("lifetime", 1))
    if interest_rate == 0:
        crf = 1 / lifetime
    else:
        crf = (((1 + interest_rate) ** lifetime) * interest_rate) / (((1 + interest_rate) ** lifetime) - 1)

    for battery_size_kwh in battery_sizes_kwh:
        print(f"Running sensitivity case for fixed battery size = {battery_size_kwh} kWh")
        if float(battery_size_kwh) == 0.0:
            solution_summary = {
                "battery_capacity_kwh": 0.0,
                "objective_total_cost": float(baseline_summary["no_battery_total_cost"]),
                "opex": float(baseline_summary["no_battery_total_cost"]),
                "import_cost": float(baseline_summary["no_battery_import_cost"]),
                "fixed_om_cost": 0.0,
                "annualized_battery_cost": 0.0,
                "peak_demand_cost": float(baseline_summary.get("no_battery_peak_demand_cost", 0.0)),
                "curtailed_energy_kwh": 0.0,
                **baseline_summary,
            }
            rows.append(
                {
                    "battery_size_kwh": 0.0,
                    "optimized_battery_capacity_kwh": 0.0,
                    "optimized_battery_power_capacity_kw": 0.0,
                    "objective_total_cost": float(solution_summary["objective_total_cost"]),
                    "import_cost": float(solution_summary["import_cost"]),
                    "peak_demand_cost": float(solution_summary["peak_demand_cost"]),
                    "annualized_battery_cost": 0.0,
                    "annual_savings": 0.0,
                    "npv": 0.0,
                    "irr": np.nan,
                    "payback_years": np.nan,
                    "discounted_payback_years": np.nan,
                    "status": "baseline",
                }
            )
            continue

        try:
            model, slacks, solution_handles = opt.setup(
                input_dict,
                debug_infeasibility=DEBUG_INFEASIBILITY,
                fixed_battery_capacity_kwh=float(battery_size_kwh),
            )
            model = opt.optimize_model(
                model,
                slacks=slacks if DEBUG_INFEASIBILITY else None,
                debug_infeasibility=DEBUG_INFEASIBILITY,
            )
            solution_summary = opt.summarize_solution(model, solution_handles)
            solution_summary.update(baseline_summary)
            financial_summary = rp.compute_financial_summary(input_dict, solution_summary)
            status = "optimal"
        except ValueError as exc:
            print(f"Sensitivity case {battery_size_kwh} kWh is infeasible: {exc}")
            rows.append(
                {
                    "battery_size_kwh": float(battery_size_kwh),
                    "optimized_battery_capacity_kwh": float(battery_size_kwh),
                    "optimized_battery_power_capacity_kw": np.nan,
                    "objective_total_cost": np.nan,
                    "import_cost": np.nan,
                    "peak_demand_cost": np.nan,
                    "annualized_battery_cost": np.nan,
                    "annual_savings": np.nan,
                    "npv": np.nan,
                    "irr": np.nan,
                    "payback_years": np.nan,
                    "discounted_payback_years": np.nan,
                    "status": "infeasible",
                }
            )
            continue

        rows.append(
            {
                "battery_size_kwh": float(battery_size_kwh),
                "optimized_battery_capacity_kwh": float(solution_summary["battery_capacity_kwh"]),
                "optimized_battery_power_capacity_kw": float(solution_summary.get("battery_power_capacity_kw", np.nan)),
                "objective_total_cost": float(solution_summary["objective_total_cost"]),
                "import_cost": float(solution_summary["import_cost"]),
                "peak_demand_cost": float(solution_summary["peak_demand_cost"]),
                "annualized_battery_cost": float(solution_summary["annualized_battery_cost"]),
                "annual_savings": float(financial_summary["annual_savings"]),
                "npv": float(financial_summary["npv"]),
                "irr": float(financial_summary["irr"]),
                "payback_years": float(financial_summary["payback_years"]),
                "discounted_payback_years": float(financial_summary["discounted_payback_years"]),
                "status": status,
            }
        )

    return pd.DataFrame(rows)


def build_default_sensitivity_sizes(input_dict: dict, optimized_capacity_kwh: float) -> list[float]:
    """
    Build a meaningful default sensitivity grid when user does not provide one.

    Strategy:
    - Always include baseline (0 kWh)
    - Sample around optimized size (50%, 75%, 100%, 125%, 150%, 200%)
    - Add a small-size point to show early economics
    - Cap values at Battery_max_capacity
    """
    cap_max = float(input_dict["parameters"].get("Battery_max_capacity", 0.0))
    if cap_max <= 0:
        return [0.0]

    optimized_capacity_kwh = float(max(0.0, optimized_capacity_kwh))
    if optimized_capacity_kwh < 1e-6:
        # If optimized is ~0, use a broad low-to-high scan.
        fractions = [0.0, 0.1, 0.25, 0.5, 0.75, 1.0]
        sizes = [cap_max * f for f in fractions]
    else:
        sizes = [
            0.0,
            min(cap_max, optimized_capacity_kwh * 0.5),
            min(cap_max, optimized_capacity_kwh * 0.75),
            min(cap_max, optimized_capacity_kwh),
            min(cap_max, optimized_capacity_kwh * 1.25),
            min(cap_max, optimized_capacity_kwh * 1.5),
            min(cap_max, optimized_capacity_kwh * 2.0),
        ]
        # Add one low absolute size point for better curve visibility.
        sizes.append(min(cap_max, max(100.0, optimized_capacity_kwh * 0.25)))

    # Unique, sorted, rounded for cleaner exports.
    sizes = sorted({round(float(s), 3) for s in sizes if s >= 0})
    return sizes


import data_preprocessing as _dpp
_dpp.refresh_config_from_excel()
importlib.reload(config)

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
solution_summary.update(rp.compute_financial_summary(input_dict, solution_summary))
save_run_artifacts(run_dir, input_dict, solution_summary, run_seconds)

battery_soc = [v.solution_value() for v in solution_handles["battery_level_vars"]]
battery_charge_power = [v.solution_value() for v in solution_handles["battery_in_flow_vars"]]
battery_discharge_power = [v.solution_value() for v in solution_handles["battery_out_flow_vars"]]
pv_flow = [v.solution_value() for v in solution_handles["pv_out_flow_vars"]]
spill_flow = [v.solution_value() for v in solution_handles["spill_flow_vars"]]
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
    spill_flow=spill_flow,
    grid_flow=grid_flow,
    total_load=total_load,
    battery_charge_power=battery_charge_power,
    battery_discharge_power=battery_discharge_power,
)

run_battery_size_sensitivity_flag = getattr(config, "run_battery_size_sensitivity", False)
battery_sensitivity_sizes_kwh = getattr(config, "battery_sensitivity_sizes_kwh", [])
if run_battery_size_sensitivity_flag and battery_sensitivity_sizes_kwh:
    sensitivity_df = run_battery_size_sensitivity(input_dict, battery_sensitivity_sizes_kwh)
    sensitivity_df.to_csv(run_dir / "battery_size_sensitivity.csv", index=False)
    print("Saved battery size sensitivity to", run_dir / "battery_size_sensitivity.csv")
elif run_battery_size_sensitivity_flag:
    auto_sizes = build_default_sensitivity_sizes(
        input_dict,
        solution_summary.get("battery_capacity_kwh", 0.0),
    )
    print(
        "No explicit battery_sensitivity_sizes_kwh provided. "
        f"Using default sizes: {auto_sizes}"
    )
    sensitivity_df = run_battery_size_sensitivity(input_dict, auto_sizes)
    sensitivity_df.to_csv(run_dir / "battery_size_sensitivity.csv", index=False)
    print("Saved battery size sensitivity to", run_dir / "battery_size_sensitivity.csv")

if getattr(config, "generate_pdf_report", True):
    settings_snapshot = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "debug_infeasibility": DEBUG_INFEASIBILITY,
        "load_existing_input_dict": config.load_existing_input_dict,
        "max_timesteps": config.max_timesteps,
        "timesteps_used": len(input_dict.get("total_demand", [])),
        "parameters": input_dict.get("parameters", {}),
    }
    report_path = report_gen.generate_pdf_report(
        run_dir=run_dir,
        solution_summary=solution_summary,
        settings_snapshot=settings_snapshot,
        input_dict=input_dict,
    )
    print("Saved PDF report to", report_path)

print("Optimization finished")
print("Objective Total Cost", solution_summary["objective_total_cost"])
print("Battery Capacity [kWh]", solution_summary["battery_capacity_kwh"])
print("Battery Power Capacity [kW]", solution_summary.get("battery_power_capacity_kw", 0.0))
print("OPEX", solution_summary["opex"])
print("Import Cost", solution_summary["import_cost"])
print("Fixed O&M Cost", solution_summary["fixed_om_cost"])
peak_type = input_dict["parameters"].get("peak_shaving_granularity", input_dict["parameters"].get("peak_shaving_frequency", "yearly"))
print("Peak Demand Type and Cost", peak_type, solution_summary["peak_demand_cost"])
if solution_summary.get("yearly_peak") is not None:
    print("Yearly Peak", solution_summary["yearly_peak"])
if solution_summary.get("monthly_peaks"):
    print("Monthly Peaks", solution_summary["monthly_peaks"])
    print("Sum Monthly Peaks", solution_summary["sum_monthly_peaks"])
print("Annualized Battery Cost", solution_summary["annualized_battery_cost"])
print("Runtime [s]", run_seconds)
print("Saved run artifacts to", run_dir)
#todo: check if data has been lost and act accordingly (dario: It might be that the PV modules were down for some reason. 
# Just copy paste two weeks before that and two weeks after that for the missing PV data (24.10. Until 7.11. for the 
# first two weeks of missing data and 10.12. Until 24.12 for the last two weeks of missing data, it’s an assumption) )
#todo: ask how flexible the data input for ev is and how smartest way to implement

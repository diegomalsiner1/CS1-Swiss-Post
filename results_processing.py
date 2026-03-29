import numpy as np
import pandas as pd
import json
from pathlib import Path

def compute_financial_summary(input_dict, solution_summary):
    lifetime = int(input_dict["parameters"].get("lifetime", 20))
    discount_rate = float(input_dict["parameters"].get("interest_rate", 0.06))
    invest_cost_per_kwh = float(input_dict["parameters"].get("Battery_invest_cost", 0.0))
    battery_capacity = float(solution_summary.get("battery_capacity_kwh", 0.0))

    investment = -battery_capacity * invest_cost_per_kwh
    no_battery_import_cost = float(solution_summary.get("no_battery_import_cost", np.nan))
    battery_import_cost = float(solution_summary.get("import_cost", np.nan))
    annual_savings = no_battery_import_cost - battery_import_cost

    years = list(range(0, lifetime + 1))
    cashflows = [investment] + [annual_savings] * lifetime
    discount_factors = [(1 + discount_rate) ** (-y) for y in years]
    discounted_cashflows = [cf * df for cf, df in zip(cashflows, discount_factors)]
    
    npv = sum(discounted_cashflows)

    # Payback period
    cumulative = 0.0
    payback = np.nan
    for year, cf in enumerate(cashflows):
        cumulative += cf
        if cumulative >= 0 and year > 0:
            payback = year
            break

    annual_financials_df = pd.DataFrame({
        "year": years,
        "cashflow": cashflows,
        "discounted_cashflow": discounted_cashflows,
    })

    return {
        "investment_cost": -investment,
        "annual_savings": annual_savings,
        "npv": npv,
        "payback_years": payback,
        "annual_financials_df": annual_financials_df,
    }

def export_results(run_dir, solution_summary, battery_soc, timestamps=None, input_dict=None, pv_flow=None, grid_flow=None, total_load=None):
    """Saves results to CSV and JSON, handling array length mismatches."""
    run_dir = Path(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)

    # 1. Handle Financials
    financial_results = {}
    if input_dict is not None:
        financial_results = compute_financial_summary(input_dict, solution_summary)
        solution_summary["npv"] = financial_results["npv"]
        solution_summary["payback_years"] = financial_results["payback_years"]
        financial_results["annual_financials_df"].to_csv(run_dir / "financial_cashflows.csv", index=False)

    # 2. Save KPI Summary to JSON
    clean_summary = {
        k: (
            {str(sub_k): sub_v for sub_k, sub_v in v.items()}
            if isinstance(v, dict) else v
        )
        for k, v in solution_summary.items()
        if isinstance(v, (int, float, str, list, dict))
    }
    with open(run_dir / "results_summary.json", "w") as f:
        json.dump(clean_summary, f, indent=4)

    # 3. Robust Time Series Alignment
    # Collect all available series
    data_map = {}
    if timestamps is not None: data_map["timestamp"] = timestamps
    if battery_soc is not None: data_map["battery_soc"] = battery_soc
    if pv_flow is not None: data_map["pv_flow"] = pv_flow
    if grid_flow is not None: data_map["grid_flow"] = grid_flow
    if total_load is not None: data_map["total_load"] = total_load

    # Find the minimum length across all provided arrays
    min_len = min(len(v) for v in data_map.values())
    
    # Truncate all arrays to the same length so pandas doesn't complain
    aligned_data = {k: list(v)[:min_len] for k, v in data_map.items()}

    ts_df = pd.DataFrame(aligned_data)
    ts_df.to_csv(run_dir / "timeseries_results.csv", index=False)

    print(f"Exported {min_len} timesteps to {run_dir}")

    return {
        "run_dir": str(run_dir),
        "summary_path": str(run_dir / "results_summary.json"),
        "timeseries_path": str(run_dir / "timeseries_results.csv")
    }
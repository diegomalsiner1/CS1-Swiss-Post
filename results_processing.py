import numpy as np
import pandas as pd
import json
from pathlib import Path

def _npv_at_rate(cashflows, rate):
    return sum(cf / ((1 + rate) ** year) for year, cf in enumerate(cashflows))


def _solve_irr(cashflows, tol=1e-7, max_iter=200):
    if not cashflows:
        return np.nan
    if not any(cf < 0 for cf in cashflows) or not any(cf > 0 for cf in cashflows):
        return np.nan

    lower = -0.9999
    upper = 0.1
    npv_lower = _npv_at_rate(cashflows, lower)
    npv_upper = _npv_at_rate(cashflows, upper)

    while npv_lower * npv_upper > 0 and upper < 1e6:
        upper = (upper + 1.0) * 2.0 - 1.0
        npv_upper = _npv_at_rate(cashflows, upper)

    if npv_lower == 0:
        return lower
    if npv_upper == 0:
        return upper
    if npv_lower * npv_upper > 0:
        return np.nan

    for _ in range(max_iter):
        mid = (lower + upper) / 2.0
        npv_mid = _npv_at_rate(cashflows, mid)
        if abs(npv_mid) < tol:
            return mid
        if npv_lower * npv_mid <= 0:
            upper = mid
            npv_upper = npv_mid
        else:
            lower = mid
            npv_lower = npv_mid

    return (lower + upper) / 2.0


def _estimate_replacement_timing(input_dict, solution_summary):
    params = input_dict["parameters"]
    cycle_life = params.get("battery_cycle_life")
    calendar_life = params.get("battery_calendar_life_years")
    fixed_replacement_year = params.get("battery_replacement_year")
    equivalent_full_cycles = float(solution_summary.get("equivalent_full_cycles", np.nan))
    battery_capacity = float(solution_summary.get("battery_capacity_kwh", 0.0))

    if battery_capacity <= 1e-9:
        return None, None, np.nan

    candidate_years = []
    cycle_based_years = np.nan

    if cycle_life is not None:
        cycle_life = float(cycle_life)
        if cycle_life <= 0:
            raise ValueError("battery_cycle_life must be positive when provided")
        if np.isfinite(equivalent_full_cycles) and equivalent_full_cycles > 1e-9:
            cycle_based_years = cycle_life / equivalent_full_cycles
            candidate_years.append(("cycle_life", cycle_based_years))

    if calendar_life is not None:
        calendar_life = float(calendar_life)
        if calendar_life <= 0:
            raise ValueError("battery_calendar_life_years must be positive when provided")
        candidate_years.append(("calendar_life", calendar_life))

    if candidate_years:
        replacement_basis, replacement_year_estimate = min(candidate_years, key=lambda x: x[1])
        replacement_year = int(np.ceil(replacement_year_estimate))
        return replacement_year, replacement_basis, cycle_based_years

    if fixed_replacement_year is not None:
        fixed_replacement_year = int(fixed_replacement_year)
        if fixed_replacement_year <= 0:
            raise ValueError("battery_replacement_year must be positive when provided")
        return fixed_replacement_year, "fixed_year", cycle_based_years

    return None, None, cycle_based_years


def compute_financial_summary(input_dict, solution_summary):
    lifetime = int(input_dict["parameters"].get("lifetime", 20))
    discount_rate = float(input_dict["parameters"].get("interest_rate", 0.06))
    invest_cost_per_kwh = float(
        input_dict["parameters"].get(
            "Battery_energy_invest_cost",
            input_dict["parameters"].get("Battery_invest_cost", 0.0),
        )
    )
    invest_cost_per_kw = float(input_dict["parameters"].get("Battery_power_invest_cost", 0.0))
    battery_capacity = float(solution_summary.get("battery_capacity_kwh", 0.0))
    battery_power_capacity = float(solution_summary.get("battery_power_capacity_kw", 0.0))
    replacement_cost_fraction = float(input_dict["parameters"].get("battery_replacement_cost_fraction", 0.0))
    replacement_year, replacement_basis, cycle_based_replacement_years = _estimate_replacement_timing(
        input_dict,
        solution_summary,
    )

    energy_investment_cost = battery_capacity * invest_cost_per_kwh
    power_investment_cost = battery_power_capacity * invest_cost_per_kw
    investment = -(energy_investment_cost + power_investment_cost)
    replacement_cost = (energy_investment_cost + power_investment_cost) * replacement_cost_fraction
    no_battery_total_cost = float(solution_summary.get("no_battery_total_cost", np.nan))
    optimized_total_cost = float(solution_summary.get("objective_total_cost", np.nan))
    annualized_battery_cost = float(solution_summary.get("annualized_battery_cost", 0.0))
    annual_total_cost_reduction = no_battery_total_cost - optimized_total_cost
    annual_savings = no_battery_total_cost - (optimized_total_cost - annualized_battery_cost)

    years = list(range(0, lifetime + 1))
    cashflows = [investment] + [annual_savings] * lifetime
    if replacement_year is not None and replacement_year <= lifetime:
        cashflows[replacement_year] -= replacement_cost
    discount_factors = [(1 + discount_rate) ** (-y) for y in years]
    discounted_cashflows = [cf * df for cf, df in zip(cashflows, discount_factors)]
    
    npv = sum(discounted_cashflows)
    irr = _solve_irr(cashflows)

    # Simple payback period
    cumulative = 0.0
    payback = np.nan
    for year, cf in enumerate(cashflows):
        cumulative += cf
        if cumulative >= 0 and year > 0:
            payback = year
            break

    # Discounted payback period
    discounted_cumulative = 0.0
    discounted_payback = np.nan
    for year, dcf in enumerate(discounted_cashflows):
        discounted_cumulative += dcf
        if discounted_cumulative >= 0 and year > 0:
            discounted_payback = year
            break

    annual_financials_df = pd.DataFrame({
        "year": years,
        "cashflow": cashflows,
        "discounted_cashflow": discounted_cashflows,
    })

    return {
        "investment_cost": -investment,
        "energy_investment_cost": energy_investment_cost,
        "power_investment_cost": power_investment_cost,
        "replacement_cost": replacement_cost,
        "replacement_year": replacement_year,
        "replacement_basis": replacement_basis,
        "cycle_based_replacement_years": cycle_based_replacement_years,
        "annual_total_cost_reduction": annual_total_cost_reduction,
        "annual_savings": annual_savings,
        "npv": npv,
        "irr": irr,
        "payback_years": payback,
        "discounted_payback_years": discounted_payback,
        "annual_financials_df": annual_financials_df,
    }


def build_baseline_vs_optimized_table(solution_summary, financial_results):
    baseline_import_cost = float(solution_summary.get("no_battery_import_cost", np.nan))
    optimized_import_cost = float(solution_summary.get("import_cost", np.nan))
    baseline_peak_cost = float(solution_summary.get("no_battery_peak_demand_cost", np.nan))
    optimized_peak_cost = float(solution_summary.get("peak_demand_cost", np.nan))
    baseline_total_cost = float(solution_summary.get("no_battery_total_cost", np.nan))
    optimized_total_cost = float(solution_summary.get("objective_total_cost", np.nan))
    investment_cost = float(financial_results.get("investment_cost", np.nan))
    energy_investment_cost = float(financial_results.get("energy_investment_cost", np.nan))
    power_investment_cost = float(financial_results.get("power_investment_cost", np.nan))
    replacement_cost = float(financial_results.get("replacement_cost", np.nan))
    replacement_year = financial_results.get("replacement_year")
    replacement_basis = financial_results.get("replacement_basis")
    annualized_battery_cost = float(solution_summary.get("annualized_battery_cost", np.nan))
    npv = float(financial_results.get("npv", np.nan))
    irr = float(financial_results.get("irr", np.nan))
    payback_years = float(financial_results.get("payback_years", np.nan))
    discounted_payback_years = float(financial_results.get("discounted_payback_years", np.nan))

    total_annual_cost_reduction = baseline_total_cost - optimized_total_cost
    import_cost_savings = baseline_import_cost - optimized_import_cost
    cost_reduction_pct = np.nan
    if np.isfinite(baseline_total_cost) and baseline_total_cost != 0:
        cost_reduction_pct = (total_annual_cost_reduction / baseline_total_cost) * 100

    comparison_rows = [
        {
            "Metric": "Battery size",
            "Baseline": 0.0,
            "Optimized": float(solution_summary.get("battery_capacity_kwh", 0.0)),
            "Optimized - Baseline": float(solution_summary.get("battery_capacity_kwh", 0.0)),
            "Unit": "kWh",
        },
        {
            "Metric": "Battery power rating",
            "Baseline": 0.0,
            "Optimized": float(solution_summary.get("battery_power_capacity_kw", 0.0)),
            "Optimized - Baseline": float(solution_summary.get("battery_power_capacity_kw", 0.0)),
            "Unit": "kW",
        },
        {
            "Metric": "Annual import cost",
            "Baseline": baseline_import_cost,
            "Optimized": optimized_import_cost,
            "Optimized - Baseline": optimized_import_cost - baseline_import_cost,
            "Unit": "CHF/year",
        },
        {
            "Metric": "Total annual cost",
            "Baseline": baseline_total_cost,
            "Optimized": optimized_total_cost,
            "Optimized - Baseline": optimized_total_cost - baseline_total_cost,
            "Unit": "CHF/year",
        },
        {
            "Metric": "Annual peak-demand cost",
            "Baseline": baseline_peak_cost,
            "Optimized": optimized_peak_cost,
            "Optimized - Baseline": optimized_peak_cost - baseline_peak_cost,
            "Unit": "CHF/year",
        },
        {
            "Metric": "Annual import cost savings",
            "Baseline": 0.0,
            "Optimized": import_cost_savings,
            "Optimized - Baseline": import_cost_savings,
            "Unit": "CHF/year",
        },
        {
            "Metric": "Annual total cost reduction",
            "Baseline": 0.0,
            "Optimized": total_annual_cost_reduction,
            "Optimized - Baseline": total_annual_cost_reduction,
            "Unit": "CHF/year",
        },
        {
            "Metric": "CAPEX",
            "Baseline": 0.0,
            "Optimized": investment_cost,
            "Optimized - Baseline": investment_cost,
            "Unit": "CHF",
        },
        {
            "Metric": "CAPEX energy component",
            "Baseline": 0.0,
            "Optimized": energy_investment_cost,
            "Optimized - Baseline": energy_investment_cost,
            "Unit": "CHF",
        },
        {
            "Metric": "CAPEX power component",
            "Baseline": 0.0,
            "Optimized": power_investment_cost,
            "Optimized - Baseline": power_investment_cost,
            "Unit": "CHF",
        },
        {
            "Metric": "Annualized battery cost",
            "Baseline": 0.0,
            "Optimized": annualized_battery_cost,
            "Optimized - Baseline": annualized_battery_cost,
            "Unit": "CHF/year",
        },
        {
            "Metric": "Replacement cost",
            "Baseline": 0.0,
            "Optimized": replacement_cost,
            "Optimized - Baseline": replacement_cost,
            "Unit": (
                f"CHF (year {replacement_year}, {replacement_basis})"
                if replacement_year is not None and replacement_basis
                else (f"CHF (year {replacement_year})" if replacement_year is not None else "CHF")
            ),
        },
        {
            "Metric": "NPV",
            "Baseline": np.nan,
            "Optimized": npv,
            "Optimized - Baseline": np.nan,
            "Unit": "CHF",
        },
        {
            "Metric": "IRR",
            "Baseline": np.nan,
            "Optimized": irr,
            "Optimized - Baseline": np.nan,
            "Unit": "-",
        },
        {
            "Metric": "Payback",
            "Baseline": np.nan,
            "Optimized": payback_years,
            "Optimized - Baseline": np.nan,
            "Unit": "years",
        },
        {
            "Metric": "Discounted payback",
            "Baseline": np.nan,
            "Optimized": discounted_payback_years,
            "Optimized - Baseline": np.nan,
            "Unit": "years",
        },
        {
            "Metric": "Cost reduction",
            "Baseline": 0.0,
            "Optimized": cost_reduction_pct,
            "Optimized - Baseline": cost_reduction_pct,
            "Unit": "%",
        },
    ]

    return pd.DataFrame(comparison_rows)


def compute_baseline_grid_import_series(input_dict):
    demand = np.asarray(input_dict.get("total_demand", []), dtype=float)
    pv_cf = np.asarray(input_dict.get("PV_capacity_factor", []), dtype=float)
    pv_max_capacity = float(input_dict["parameters"]["PV_max_capacity"])
    surplus_handling = input_dict["parameters"].get("surplus_handling", "curtail").lower()

    if len(demand) != len(pv_cf):
        raise ValueError("Input timeseries lengths must match for baseline grid import computation.")

    pv_available = pv_cf * pv_max_capacity
    if surplus_handling == "must_absorb" and np.any(pv_available > demand + 1e-9):
        raise ValueError(
            "No-battery baseline grid import is infeasible under surplus_handling='must_absorb' because PV exceeds demand."
        )
    return np.maximum(demand - pv_available, 0.0)


def build_peak_metrics_tables(timestamps, baseline_grid_import, optimized_grid_import, top_n=10):
    peak_df = pd.DataFrame({
        "timestamp": pd.to_datetime(list(timestamps)),
        "baseline_grid_import": np.asarray(baseline_grid_import, dtype=float),
        "optimized_grid_import": np.asarray(optimized_grid_import, dtype=float),
    })
    peak_df["peak_reduction"] = peak_df["baseline_grid_import"] - peak_df["optimized_grid_import"]

    max_before = float(peak_df["baseline_grid_import"].max())
    max_after = float(peak_df["optimized_grid_import"].max())
    p95_before = float(peak_df["baseline_grid_import"].quantile(0.95))
    p95_after = float(peak_df["optimized_grid_import"].quantile(0.95))

    max_reduction_pct = np.nan
    if max_before != 0:
        max_reduction_pct = ((max_before - max_after) / max_before) * 100

    p95_reduction_pct = np.nan
    if p95_before != 0:
        p95_reduction_pct = ((p95_before - p95_after) / p95_before) * 100

    summary_rows = [
        {
            "Metric": "Maximum grid import",
            "Before battery": max_before,
            "After battery": max_after,
            "Reduction": max_before - max_after,
            "Reduction %": max_reduction_pct,
            "Unit": "kW",
        },
        {
            "Metric": "95th percentile grid import",
            "Before battery": p95_before,
            "After battery": p95_after,
            "Reduction": p95_before - p95_after,
            "Reduction %": p95_reduction_pct,
            "Unit": "kW",
        },
    ]

    top_peak_intervals = peak_df.sort_values("peak_reduction", ascending=False).head(top_n).copy()
    return pd.DataFrame(summary_rows), top_peak_intervals


def build_monthly_summary_table(timestamps, baseline_grid_import, optimized_grid_import, electricity_price):
    monthly_df = pd.DataFrame({
        "timestamp": pd.to_datetime(list(timestamps)),
        "baseline_grid_import": np.asarray(baseline_grid_import, dtype=float),
        "optimized_grid_import": np.asarray(optimized_grid_import, dtype=float),
        "electricity_price": np.asarray(electricity_price, dtype=float),
    })
    monthly_df["month"] = monthly_df["timestamp"].dt.to_period("M").astype(str)
    timestep_hours = 0.25
    monthly_df["baseline_import_cost"] = (
        monthly_df["baseline_grid_import"] * monthly_df["electricity_price"] * timestep_hours
    )
    monthly_df["optimized_import_cost"] = (
        monthly_df["optimized_grid_import"] * monthly_df["electricity_price"] * timestep_hours
    )

    summary_df = (
        monthly_df.groupby("month", as_index=False)
        .agg(
            monthly_import_cost_before=("baseline_import_cost", "sum"),
            monthly_import_cost_after=("optimized_import_cost", "sum"),
            monthly_peak_before=("baseline_grid_import", "max"),
            monthly_peak_after=("optimized_grid_import", "max"),
        )
    )
    summary_df["monthly_savings"] = (
        summary_df["monthly_import_cost_before"] - summary_df["monthly_import_cost_after"]
    )
    summary_df["monthly_peak_reduction"] = (
        summary_df["monthly_peak_before"] - summary_df["monthly_peak_after"]
    )

    return summary_df[
        [
            "month",
            "monthly_import_cost_before",
            "monthly_import_cost_after",
            "monthly_savings",
            "monthly_peak_before",
            "monthly_peak_after",
            "monthly_peak_reduction",
        ]
    ]


def build_battery_utilization_table(solution_summary, battery_soc, battery_charge_power, battery_discharge_power):
    timestep_hours = 0.25
    battery_soc = np.asarray(battery_soc, dtype=float)
    battery_charge_power = np.asarray(battery_charge_power, dtype=float)
    battery_discharge_power = np.asarray(battery_discharge_power, dtype=float)

    installed_capacity = float(solution_summary.get("battery_capacity_kwh", 0.0))
    charged_energy = float(battery_charge_power.sum() * timestep_hours)
    discharged_energy = float(battery_discharge_power.sum() * timestep_hours)
    equivalent_cycles = np.nan
    if installed_capacity > 1e-9:
        equivalent_cycles = discharged_energy / installed_capacity

    avg_soc = float(battery_soc.mean()) if battery_soc.size else np.nan
    min_soc = float(battery_soc.min()) if battery_soc.size else np.nan
    max_soc = float(battery_soc.max()) if battery_soc.size else np.nan

    soc_empty_threshold = 0.10 * installed_capacity
    soc_full_threshold = 0.90 * installed_capacity
    hours_near_empty = float((battery_soc <= soc_empty_threshold).sum() * timestep_hours)
    hours_near_full = float((battery_soc >= soc_full_threshold).sum() * timestep_hours)

    summary_rows = [
        {"Metric": "Charged energy", "Value": charged_energy, "Unit": "kWh/year"},
        {"Metric": "Discharged energy", "Value": discharged_energy, "Unit": "kWh/year"},
        {"Metric": "Equivalent full cycles", "Value": equivalent_cycles, "Unit": "cycles/year"},
        {"Metric": "Average state of charge", "Value": avg_soc, "Unit": "kWh"},
        {"Metric": "Minimum state of charge", "Value": min_soc, "Unit": "kWh"},
        {"Metric": "Maximum state of charge", "Value": max_soc, "Unit": "kWh"},
        {"Metric": "Hours near empty (<=10%)", "Value": hours_near_empty, "Unit": "hours"},
        {"Metric": "Hours near full (>=90%)", "Value": hours_near_full, "Unit": "hours"},
        {
            "Metric": "Maximum charge power used",
            "Value": float(battery_charge_power.max()) if battery_charge_power.size else np.nan,
            "Unit": "kW",
        },
        {
            "Metric": "Maximum discharge power used",
            "Value": float(battery_discharge_power.max()) if battery_discharge_power.size else np.nan,
            "Unit": "kW",
        },
    ]

    return pd.DataFrame(summary_rows)


def export_results(
    run_dir,
    solution_summary,
    battery_soc,
    timestamps=None,
    input_dict=None,
    pv_flow=None,
    spill_flow=None,
    grid_flow=None,
    total_load=None,
    battery_charge_power=None,
    battery_discharge_power=None,
):
    """Saves results to CSV and JSON, handling array length mismatches."""
    run_dir = Path(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)

    # 1. Handle Financials
    financial_results = {}
    if input_dict is not None:
        financial_results = compute_financial_summary(input_dict, solution_summary)
        solution_summary["npv"] = financial_results["npv"]
        solution_summary["irr"] = financial_results["irr"]
        solution_summary["payback_years"] = financial_results["payback_years"]
        solution_summary["discounted_payback_years"] = financial_results["discounted_payback_years"]
        financial_results["annual_financials_df"].to_csv(run_dir / "financial_cashflows.csv", index=False)
        comparison_df = build_baseline_vs_optimized_table(solution_summary, financial_results)
        comparison_df.to_csv(run_dir / "baseline_vs_optimized.csv", index=False)

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
    if spill_flow is not None: data_map["spill_flow"] = spill_flow
    if grid_flow is not None: data_map["grid_flow"] = grid_flow
    if total_load is not None: data_map["total_load"] = total_load
    if battery_charge_power is not None: data_map["battery_charge_power"] = battery_charge_power
    if battery_discharge_power is not None: data_map["battery_discharge_power"] = battery_discharge_power
    if input_dict is not None:
        data_map["baseline_grid_import"] = compute_baseline_grid_import_series(input_dict)
        if "electricity_price" in input_dict:
            data_map["electricity_price"] = input_dict["electricity_price"]

    # Find the minimum length across all provided arrays
    min_len = min(len(v) for v in data_map.values())
    
    # Truncate all arrays to the same length so pandas doesn't complain
    aligned_data = {k: list(v)[:min_len] for k, v in data_map.items()}
    if "baseline_grid_import" in aligned_data and "grid_flow" in aligned_data:
        aligned_data["peak_reduction"] = (
            np.asarray(aligned_data["baseline_grid_import"], dtype=float)
            - np.asarray(aligned_data["grid_flow"], dtype=float)
        ).tolist()

    ts_df = pd.DataFrame(aligned_data)
    ts_df.to_csv(run_dir / "timeseries_results.csv", index=False)

    peak_metrics_path = None
    top_peak_path = None
    monthly_summary_path = None
    battery_utilization_path = None
    if "timestamp" in ts_df.columns and "baseline_grid_import" in ts_df.columns and "grid_flow" in ts_df.columns:
        peak_metrics_df, top_peak_intervals_df = build_peak_metrics_tables(
            ts_df["timestamp"],
            ts_df["baseline_grid_import"],
            ts_df["grid_flow"],
        )
        peak_metrics_path = run_dir / "peak_metrics.csv"
        top_peak_path = run_dir / "top_peak_intervals.csv"
        peak_metrics_df.to_csv(peak_metrics_path, index=False)
        top_peak_intervals_df.to_csv(top_peak_path, index=False)
        if input_dict is not None:
            monthly_summary_df = build_monthly_summary_table(
                ts_df["timestamp"],
                ts_df["baseline_grid_import"],
                ts_df["grid_flow"],
                input_dict["electricity_price"][:min_len],
            )
            monthly_summary_path = run_dir / "monthly_summary.csv"
            monthly_summary_df.to_csv(monthly_summary_path, index=False)
    if (
        "battery_soc" in ts_df.columns
        and "battery_charge_power" in ts_df.columns
        and "battery_discharge_power" in ts_df.columns
    ):
        battery_utilization_df = build_battery_utilization_table(
            solution_summary,
            ts_df["battery_soc"],
            ts_df["battery_charge_power"],
            ts_df["battery_discharge_power"],
        )
        battery_utilization_path = run_dir / "battery_utilization_summary.csv"
        battery_utilization_df.to_csv(battery_utilization_path, index=False)

    print(f"Exported {min_len} timesteps to {run_dir}")

    return {
        "run_dir": str(run_dir),
        "summary_path": str(run_dir / "results_summary.json"),
        "timeseries_path": str(run_dir / "timeseries_results.csv"),
        "comparison_path": str(run_dir / "baseline_vs_optimized.csv") if input_dict is not None else None,
        "peak_metrics_path": str(peak_metrics_path) if peak_metrics_path is not None else None,
        "top_peak_intervals_path": str(top_peak_path) if top_peak_path is not None else None,
        "monthly_summary_path": str(monthly_summary_path) if monthly_summary_path is not None else None,
        "battery_utilization_path": str(battery_utilization_path) if battery_utilization_path is not None else None,
    }

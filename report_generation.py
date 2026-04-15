from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path
import textwrap

# Ensure matplotlib uses a writable config/cache directory and headless backend.
_MPL_CACHE_DIR = Path(__file__).resolve().parent / ".matplotlib-cache"
_MPL_CACHE_DIR.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", str(_MPL_CACHE_DIR))

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

_BRAND = {
    "primary": "#D52B1E",
    "secondary": "#003B5C",
    "accent": "#FFB81C",
    "ok": "#2E7D32",
    "neutral": "#4E5D6C",
    "grid": "#D6DCE3",
    "header_bg": "#EEF2F7",
}

plt.rcParams.update(
    {
        "font.size": 10,
        "axes.titlesize": 14,
        "axes.labelsize": 11,
        "axes.titleweight": "bold",
        "axes.edgecolor": "#2A2A2A",
        "axes.linewidth": 0.8,
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "grid.color": _BRAND["grid"],
        "grid.linewidth": 0.8,
        "legend.frameon": False,
    }
)


def _fmt_num(value, digits: int = 2) -> str:
    try:
        if value is None or pd.isna(value):
            return "n/a"
        return f"{float(value):,.{digits}f}"
    except Exception:
        return str(value)


def _style_axes(ax, title: str, xlabel: str | None = None, ylabel: str | None = None) -> None:
    ax.set_title(title, loc="left", color=_BRAND["secondary"], pad=8)
    if xlabel:
        ax.set_xlabel(xlabel)
    if ylabel:
        ax.set_ylabel(ylabel)
    ax.grid(alpha=0.45)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def _safe_text(value: object, width: int = 22, max_lines: int = 2) -> str:
    wrapped = textwrap.wrap(str(value), width=width) or [""]
    if len(wrapped) > max_lines:
        wrapped = wrapped[:max_lines]
        wrapped[-1] = wrapped[-1][: max(0, width - 3)] + "..."
    return "\n".join(wrapped)


def _apply_external_legend(ax, ncol: int = 2, y_anchor: float = 1.16, fontsize: float = 8.5) -> None:
    handles, labels = ax.get_legend_handles_labels()
    if not handles:
        return
    ax.legend(
        handles,
        labels,
        loc="upper left",
        ncol=ncol,
        fontsize=fontsize,
        frameon=True,
        facecolor="white",
        edgecolor=_BRAND["grid"],
        framealpha=0.95,
    )


def _add_text_page(pdf: PdfPages, title: str, lines: list[str], subtitle: str | None = None) -> None:
    fig = plt.figure(figsize=(11.69, 8.27))  # A4 landscape
    fig.suptitle(title, fontsize=19, fontweight="bold", y=0.965, color=_BRAND["secondary"])
    if subtitle:
        fig.text(0.05, 0.905, subtitle, fontsize=11, color=_BRAND["neutral"])
    y = 0.865
    for line in lines:
        wrapped = textwrap.wrap(str(line), width=115) or [""]
        for wrapped_line in wrapped:
            fig.text(0.05, y, wrapped_line, fontsize=11)
            y -= 0.028
            if y < 0.05:
                pdf.savefig(fig, bbox_inches="tight")
                plt.close(fig)
                fig = plt.figure(figsize=(11.69, 8.27))
                fig.suptitle(title + " (cont.)", fontsize=16, fontweight="bold", y=0.965, color=_BRAND["secondary"])
                y = 0.90
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _add_table_page(
    pdf: PdfPages,
    title: str,
    df: pd.DataFrame,
    max_rows: int = 24,
    subtitle: str | None = None,
) -> None:
    if df is None or df.empty:
        _add_text_page(pdf, title, ["No data available for this section."], subtitle=subtitle)
        return

    start = 0
    page_idx = 1
    while start < len(df):
        chunk = df.iloc[start : start + max_rows].copy()
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.axis("off")
        chunk_title = title if len(df) <= max_rows else f"{title} (part {page_idx})"
        ax.set_title(chunk_title, fontsize=14, pad=10, loc="left", color=_BRAND["secondary"])
        if subtitle and page_idx == 1:
            fig.text(0.05, 0.93, subtitle, fontsize=10, color=_BRAND["neutral"])

        # Convert to readable strings
        display_chunk = chunk.copy()
        for col in display_chunk.columns:
            col_name = str(col)
            if len(col_name) > 24:
                col_name = "\n".join(textwrap.wrap(col_name, width=24))
            if col_name != col:
                display_chunk.rename(columns={col: col_name}, inplace=True)
        for col in display_chunk.columns:
            display_chunk[col] = display_chunk[col].apply(
                lambda x: _fmt_num(x, 2) if isinstance(x, (int, float)) else str(x)
            )

        tbl = ax.table(
            cellText=display_chunk.values,
            colLabels=display_chunk.columns,
            cellLoc="center",
            loc="center",
        )
        tbl.auto_set_font_size(False)
        n_cols = max(1, len(display_chunk.columns))
        tbl.set_fontsize(9 if n_cols <= 4 else 8 if n_cols <= 6 else 7)
        tbl.scale(1, 1.1 if n_cols > 6 else 1.2)
        for (row, col), cell in tbl.get_celld().items():
            if row == 0:
                cell.set_facecolor(_BRAND["header_bg"])
                cell.set_text_props(weight="bold", color=_BRAND["secondary"])
            elif row % 2 == 0:
                cell.set_facecolor("#FAFCFF")
            if col == 0:
                cell._loc = "left"
        fig.subplots_adjust(left=0.03, right=0.97, top=0.89, bottom=0.04)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        start += max_rows
        page_idx += 1


def _plot_grid_import_timeseries(pdf: PdfPages, df_ts: pd.DataFrame) -> None:
    if not {"timestamp", "grid_flow", "baseline_grid_import"}.issubset(df_ts.columns):
        return
    ts = df_ts[["timestamp", "baseline_grid_import", "grid_flow"]].copy().sort_values("timestamp")
    ts["timestamp"] = pd.to_datetime(ts["timestamp"], errors="coerce")
    ts = ts.dropna(subset=["timestamp"])
    if ts.empty:
        return
    daily = (
        ts.set_index("timestamp")
        .resample("D")
        .agg(
            baseline_mean=("baseline_grid_import", "mean"),
            optimized_mean=("grid_flow", "mean"),
            baseline_peak=("baseline_grid_import", "max"),
            optimized_peak=("grid_flow", "max"),
        )
        .dropna(how="all")
        .reset_index()
    )
    if daily.empty:
        return
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.plot(daily["timestamp"], daily["baseline_mean"], label="Baseline daily mean", linewidth=1.3, color=_BRAND["neutral"])
    ax.plot(daily["timestamp"], daily["optimized_mean"], label="Optimized daily mean", linewidth=1.3, color=_BRAND["primary"])
    _style_axes(ax, "Grid Import Trend (Daily Mean)", xlabel="Date", ylabel="Power [kW]")
    _apply_external_legend(ax, ncol=2, y_anchor=1.20)
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.plot(daily["timestamp"], daily["baseline_peak"], label="Baseline daily peak", linewidth=1.3, color="#7A8896")
    ax.plot(daily["timestamp"], daily["optimized_peak"], label="Optimized daily peak", linewidth=1.3, color="#8B1A10")
    _style_axes(ax, "Grid Import Trend (Daily Peak)", xlabel="Date", ylabel="Power [kW]")
    _apply_external_legend(ax, ncol=2, y_anchor=1.20)
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _plot_representative_peak_days(pdf: PdfPages, df_ts: pd.DataFrame) -> None:
    required = {"timestamp", "grid_flow", "baseline_grid_import"}
    if not required.issubset(df_ts.columns):
        return
    ts = df_ts[["timestamp", "grid_flow", "baseline_grid_import"]].copy()
    ts["timestamp"] = pd.to_datetime(ts["timestamp"], errors="coerce")
    ts = ts.dropna(subset=["timestamp"]).sort_values("timestamp")
    if ts.empty:
        return

    ts["date"] = ts["timestamp"].dt.date
    daily_peak = ts.groupby("date")["baseline_grid_import"].max().sort_values(ascending=False)
    if daily_peak.empty:
        return

    selected_dates = [daily_peak.index[0]]
    median_idx = int(len(daily_peak) * 0.5)
    selected_dates.append(daily_peak.index[min(median_idx, len(daily_peak) - 1)])
    selected_dates = list(dict.fromkeys(selected_dates))

    for i, day in enumerate(selected_dates):
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        day_df = ts[ts["date"] == day]
        if day_df.empty:
            plt.close(fig)
            continue
        ax.plot(day_df["timestamp"], day_df["baseline_grid_import"], label="Baseline", color=_BRAND["neutral"], linewidth=1.4)
        ax.plot(day_df["timestamp"], day_df["grid_flow"], label="Optimized", color=_BRAND["primary"], linewidth=1.4)
        title_suffix = "highest-peak day" if i == 0 else "median-peak day"
        _style_axes(ax, f"Representative Day Dispatch ({title_suffix})", xlabel="Time", ylabel="Grid import [kW]")
        _apply_external_legend(ax, ncol=2, y_anchor=1.20)
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def _plot_dispatch_and_soc(pdf: PdfPages, df_ts: pd.DataFrame) -> None:
    required = {"timestamp", "grid_flow", "pv_flow", "total_load", "battery_charge_power", "battery_discharge_power"}
    if not required.issubset(df_ts.columns):
        return
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.plot(df_ts["timestamp"], df_ts["total_load"], label="Net load", color="#1B1B1B", linewidth=1.3)
    ax.plot(df_ts["timestamp"], df_ts["grid_flow"], label="Grid import", color=_BRAND["primary"], linewidth=1.2)
    ax.plot(df_ts["timestamp"], df_ts["pv_flow"], label="PV used", color=_BRAND["accent"], linewidth=1.2)
    ax.plot(df_ts["timestamp"], df_ts["battery_discharge_power"], label="Battery discharge", color=_BRAND["ok"], linewidth=1.1)
    ax.plot(df_ts["timestamp"], -df_ts["battery_charge_power"], label="Battery charge (negative)", color="#2563EB", linewidth=1.1)
    _style_axes(ax, "Dispatch Overview", xlabel="Timestamp", ylabel="Power [kW]")
    _apply_external_legend(ax, ncol=3, y_anchor=1.20, fontsize=8.2)
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    if "battery_soc" in df_ts.columns:
        ax.plot(df_ts["timestamp"], df_ts["battery_soc"], color=_BRAND["ok"], linewidth=1.3)
        ax.fill_between(df_ts["timestamp"], 0, df_ts["battery_soc"], alpha=0.12, color=_BRAND["ok"])
    _style_axes(ax, "Battery State of Charge", xlabel="Timestamp", ylabel="SOC [kWh]")
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _plot_representative_dispatch_weeks(pdf: PdfPages, df_ts: pd.DataFrame) -> None:
    required = {"timestamp", "grid_flow", "pv_flow", "total_load", "battery_charge_power", "battery_discharge_power"}
    if not required.issubset(df_ts.columns):
        return
    ts = df_ts.copy()
    ts["timestamp"] = pd.to_datetime(ts["timestamp"], errors="coerce")
    ts = ts.dropna(subset=["timestamp"]).sort_values("timestamp")
    if ts.empty:
        return
    ts["week"] = ts["timestamp"].dt.to_period("W").apply(lambda p: p.start_time)
    weekly_peak = ts.groupby("week")["grid_flow"].max().sort_values(ascending=False)
    if weekly_peak.empty:
        return

    selected_weeks = [weekly_peak.index[0], weekly_peak.index[min(len(weekly_peak) - 1, len(weekly_peak) // 2)]]
    selected_weeks = list(dict.fromkeys(selected_weeks))
    for i, wk in enumerate(selected_weeks):
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        wk_end = wk + pd.Timedelta(days=7)
        wk_df = ts[(ts["timestamp"] >= wk) & (ts["timestamp"] < wk_end)].copy()
        if wk_df.empty:
            plt.close(fig)
            continue
        actual_load = wk_df["grid_flow"] + wk_df["pv_flow"] + wk_df["battery_discharge_power"]
        ax.plot(wk_df["timestamp"], actual_load, color="#111111", linewidth=1.25, label="Actual load")
        ax.plot(wk_df["timestamp"], wk_df["total_load"], color=_BRAND["neutral"], linewidth=1.1, label="Net load")
        ax.plot(wk_df["timestamp"], wk_df["grid_flow"], color=_BRAND["primary"], linewidth=1.1, label="Grid import")
        ax.plot(wk_df["timestamp"], wk_df["battery_discharge_power"], color=_BRAND["ok"], linewidth=1.0, label="Battery discharge")
        ax.plot(wk_df["timestamp"], -wk_df["battery_charge_power"], color="#2563EB", linewidth=1.0, label="Battery charge (negative)")
        label = "highest peak week" if i == 0 else "median peak week"
        _style_axes(ax, f"Representative Dispatch Week ({label})", xlabel="Time", ylabel="Power [kW]")
        _apply_external_legend(ax, ncol=3, y_anchor=1.20, fontsize=8)
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def _plot_monthly_savings(pdf: PdfPages, df_monthly: pd.DataFrame) -> None:
    if df_monthly.empty or "month" not in df_monthly.columns:
        return
    if "monthly_savings" in df_monthly.columns:
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.bar(df_monthly["month"], df_monthly["monthly_savings"], color=_BRAND["secondary"])
        _style_axes(ax, "Monthly Cost Savings", xlabel="Month", ylabel="CHF/month")
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)
    if "monthly_peak_reduction" in df_monthly.columns:
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.bar(df_monthly["month"], df_monthly["monthly_peak_reduction"], color=_BRAND["accent"])
        _style_axes(ax, "Monthly Peak Reduction", xlabel="Month", ylabel="kW")
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def _plot_weekly_savings(pdf: PdfPages, df_weekly: pd.DataFrame) -> None:
    if df_weekly.empty or "week_start" not in df_weekly.columns:
        return
    weekly = df_weekly.copy()
    weekly["week_start"] = pd.to_datetime(weekly["week_start"], errors="coerce")
    weekly = weekly.dropna(subset=["week_start"]).sort_values("week_start")
    if weekly.empty:
        return
    tick_positions = weekly["week_start"].iloc[:: max(1, len(weekly) // 10)]
    if "weekly_cost_savings" in df_weekly.columns:
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.bar(weekly["week_start"], weekly["weekly_cost_savings"], color=_BRAND["secondary"])
        _style_axes(ax, "Weekly Cost Savings", xlabel="Week start", ylabel="CHF/week")
        ax.set_xticks(tick_positions)
        ax.tick_params(axis="x", rotation=35)
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)
    if "weekly_peak_reduction" in df_weekly.columns:
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        ax.bar(weekly["week_start"], weekly["weekly_peak_reduction"], color=_BRAND["accent"])
        _style_axes(ax, "Weekly Peak Reduction", xlabel="Week start", ylabel="kW")
        ax.set_xticks(tick_positions)
        ax.tick_params(axis="x", rotation=35)
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)


def _plot_cashflows(pdf: PdfPages, df_fin: pd.DataFrame) -> None:
    if df_fin.empty or not {"year", "cashflow", "discounted_cashflow"}.issubset(df_fin.columns):
        return
    fig, ax1 = plt.subplots(figsize=(11.69, 8.27))
    ax1.bar(df_fin["year"], df_fin["cashflow"], color=_BRAND["secondary"], alpha=0.75, label="Cashflow")
    _style_axes(ax1, "Financial Cashflows", xlabel="Year", ylabel="Cashflow [CHF]")

    ax2 = ax1.twinx()
    ax2.plot(
        df_fin["year"],
        df_fin["discounted_cashflow"].cumsum(),
        color=_BRAND["primary"],
        marker="o",
        label="Cumulative discounted cashflow",
    )
    ax2.set_ylabel("Cumulative discounted cashflow [CHF]")
    ax2.spines["top"].set_visible(False)

    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(
        h1 + h2,
        l1 + l2,
        loc="upper left",
        ncol=2,
        fontsize=8.5,
        frameon=True,
        facecolor="white",
        edgecolor=_BRAND["grid"],
        framealpha=0.95,
    )
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _plot_duration_curve(pdf: PdfPages, df_ts: pd.DataFrame) -> None:
    if not {"grid_flow", "baseline_grid_import"}.issubset(df_ts.columns):
        return
    baseline = df_ts["baseline_grid_import"].sort_values(ascending=False).reset_index(drop=True)
    optimized = df_ts["grid_flow"].sort_values(ascending=False).reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.plot(baseline.index, baseline, label="Baseline", color=_BRAND["neutral"], linewidth=1.4)
    ax.plot(optimized.index, optimized, label="Optimized", color=_BRAND["primary"], linewidth=1.4)
    _style_axes(ax, "Grid Import Duration Curve", xlabel="Sorted timestep index", ylabel="Grid import [kW]")
    _apply_external_legend(ax, ncol=2, y_anchor=1.20)
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _plot_sensitivity(pdf: PdfPages, df_sens: pd.DataFrame) -> None:
    if df_sens.empty or "battery_size_kwh" not in df_sens.columns:
        return
    fig, ax1 = plt.subplots(figsize=(11.69, 8.27))
    feasible_df = (
        df_sens[df_sens["status"].isin(["baseline", "optimal"])]
        if "status" in df_sens.columns
        else df_sens.copy()
    )
    infeasible_df = df_sens[df_sens["status"] == "infeasible"] if "status" in df_sens.columns else pd.DataFrame()
    feasible_df = feasible_df.sort_values("battery_size_kwh")

    if "objective_total_cost" in feasible_df.columns and not feasible_df.empty:
        ax1.plot(
            feasible_df["battery_size_kwh"],
            feasible_df["objective_total_cost"],
            marker="o",
            color=_BRAND["secondary"],
            label="TAC (feasible)",
        )
    if "objective_total_cost" in infeasible_df.columns and not infeasible_df.empty:
        infeasible_plot = infeasible_df.dropna(subset=["objective_total_cost"])
        if not infeasible_plot.empty:
            ax1.scatter(
                infeasible_plot["battery_size_kwh"],
                infeasible_plot["objective_total_cost"],
                color=_BRAND["primary"],
                marker="x",
                s=60,
                label="Infeasible points",
            )
    _style_axes(ax1, "Battery Size Sensitivity", xlabel="Battery size [kWh]", ylabel="TAC [CHF/year]")

    if "npv" in df_sens.columns:
        ax2 = ax1.twinx()
        if not feasible_df.empty:
            ax2.plot(
                feasible_df["battery_size_kwh"],
                feasible_df["npv"],
                marker="s",
                color=_BRAND["ok"],
                label="NPV (feasible)",
            )
        if not infeasible_df.empty:
            infeasible_npv = infeasible_df.dropna(subset=["npv"])
            if not infeasible_npv.empty:
                ax2.scatter(
                    infeasible_npv["battery_size_kwh"],
                    infeasible_npv["npv"],
                    color=_BRAND["primary"],
                    marker="x",
                    s=60,
                    label="NPV infeasible points",
                )
        ax2.set_ylabel("NPV [CHF]")
        h1, l1 = ax1.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax1.legend(
            h1 + h2,
            l1 + l2,
            loc="upper left",
            ncol=2,
            fontsize=8.5,
            frameon=True,
            facecolor="white",
            edgecolor=_BRAND["grid"],
            framealpha=0.95,
        )
    else:
        _apply_external_legend(ax1, ncol=2, y_anchor=1.20)
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def _build_kpi_table(solution_summary: dict) -> pd.DataFrame:
    ordered_metrics = [
        ("objective_total_cost", "TAC / Objective Total Cost [CHF/year]"),
        ("battery_capacity_kwh", "Battery Capacity [kWh]"),
        ("battery_power_capacity_kw", "Battery Power Capacity [kW]"),
        ("import_cost", "Import Cost [CHF/year]"),
        ("peak_demand_cost", "Peak Demand Cost [CHF/year]"),
        ("opex", "OPEX [CHF/year]"),
        ("annualized_battery_cost", "Annualized Battery Cost [CHF/year]"),
        ("no_battery_total_cost", "No-Battery Total Cost [CHF/year]"),
        ("no_battery_import_cost", "No-Battery Import Cost [CHF/year]"),
        ("no_battery_peak_demand_cost", "No-Battery Peak Demand Cost [CHF/year]"),
        ("npv", "NPV [CHF]"),
        ("irr", "IRR [-]"),
        ("payback_years", "Payback [years]"),
        ("discounted_payback_years", "Discounted Payback [years]"),
        ("replacement_cost", "Replacement Cost [CHF]"),
        ("replacement_year", "Replacement Year"),
        ("curtailed_energy_kwh", "Curtailed Energy [kWh/year]"),
        ("discharged_energy_kwh", "Discharged Energy [kWh/year]"),
        ("equivalent_full_cycles", "Equivalent Full Cycles [cycles/year]"),
        ("runtime_seconds", "Runtime [s]"),
    ]
    rows = []
    for key, label in ordered_metrics:
        if key in solution_summary:
            rows.append({"Metric": label, "Value": solution_summary.get(key)})
    return pd.DataFrame(rows)


def _build_settings_table(settings_snapshot: dict, input_dict: dict) -> pd.DataFrame:
    rows = [
        {"Section": "Run", "Parameter": "timestamp", "Value": settings_snapshot.get("timestamp")},
        {"Section": "Run", "Parameter": "timesteps_used", "Value": settings_snapshot.get("timesteps_used")},
        {"Section": "Run", "Parameter": "load_existing_input_dict", "Value": settings_snapshot.get("load_existing_input_dict")},
        {"Section": "Run", "Parameter": "max_timesteps", "Value": settings_snapshot.get("max_timesteps")},
    ]
    params = input_dict.get("parameters", {})
    for key in sorted(params.keys()):
        rows.append({"Section": "Model Parameter", "Parameter": key, "Value": params[key]})
    return pd.DataFrame(rows)


def _to_report_table(df: pd.DataFrame, keep_columns: list[str], rename_map: dict[str, str] | None = None, decimals: int = 2) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    available = [c for c in keep_columns if c in df.columns]
    if not available:
        return pd.DataFrame()
    out = df[available].copy()
    if rename_map:
        out.rename(columns=rename_map, inplace=True)
    for col in out.columns:
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].round(decimals)
    return out


def _build_weekly_highlights(df_weekly: pd.DataFrame) -> pd.DataFrame:
    if df_weekly is None or df_weekly.empty:
        return pd.DataFrame()
    required = {"week_start", "weekly_cost_savings", "weekly_peak_reduction"}
    if not required.issubset(df_weekly.columns):
        return pd.DataFrame()
    tmp = df_weekly.copy()
    tmp["week_start"] = pd.to_datetime(tmp["week_start"], errors="coerce")
    tmp = tmp.dropna(subset=["week_start"])
    if tmp.empty:
        return pd.DataFrame()
    best = tmp.nlargest(5, "weekly_cost_savings")[["week_start", "weekly_cost_savings", "weekly_peak_reduction"]].copy()
    best["Category"] = "Top savings week"
    worst = tmp.nsmallest(5, "weekly_cost_savings")[["week_start", "weekly_cost_savings", "weekly_peak_reduction"]].copy()
    worst["Category"] = "Lowest savings week"
    out = pd.concat([best, worst], ignore_index=True).sort_values(["Category", "week_start"])
    out["week_start"] = out["week_start"].dt.strftime("%Y-%m-%d")
    return out[["Category", "week_start", "weekly_cost_savings", "weekly_peak_reduction"]]


def _build_executive_insights(solution_summary: dict) -> list[str]:
    tac = solution_summary.get("objective_total_cost")
    baseline = solution_summary.get("no_battery_total_cost")
    npv = solution_summary.get("npv")
    payback = solution_summary.get("payback_years")
    battery_kwh = solution_summary.get("battery_capacity_kwh")
    battery_kw = solution_summary.get("battery_power_capacity_kw")

    annual_saving = None
    if tac is not None and baseline is not None:
        try:
            annual_saving = float(baseline) - float(tac)
        except Exception:
            annual_saving = None

    lines = [
        "Decision-ready highlights",
        f"- Proposed BESS size: {_fmt_num(battery_kwh)} kWh / {_fmt_num(battery_kw)} kW.",
        f"- Estimated TAC: {_fmt_num(tac)} CHF/year.",
        f"- Baseline annual cost without battery: {_fmt_num(baseline)} CHF/year.",
    ]
    if annual_saving is not None:
        lines.append(f"- Estimated annual cost reduction: {_fmt_num(annual_saving)} CHF/year.")
    lines.extend(
        [
            f"- Project NPV: {_fmt_num(npv)} CHF.",
            f"- Simple payback: {_fmt_num(payback)} years.",
            "",
            "Interpretation guidance",
            "- Positive NPV and acceptable payback generally indicate business viability.",
            "- Check peak-related charts/tables to validate grid-constraint risk reduction.",
            "- Use sensitivity pages to assess robustness under alternative battery sizes.",
        ]
    )
    return lines


def _get_recommendation(solution_summary: dict) -> tuple[str, str, str]:
    npv = solution_summary.get("npv")
    payback = solution_summary.get("payback_years")
    irr = solution_summary.get("irr")

    try:
        npv_v = float(npv) if npv is not None else None
    except Exception:
        npv_v = None
    try:
        payback_v = float(payback) if payback is not None else None
    except Exception:
        payback_v = None
    try:
        irr_v = float(irr) if irr is not None else None
    except Exception:
        irr_v = None

    if npv_v is None:
        return ("REVIEW NEEDED", _BRAND["accent"], "NPV unavailable; confirm inputs and rerun.")

    if npv_v > 0 and (payback_v is None or payback_v <= 8):
        reason = "Positive NPV with acceptable payback profile."
        if irr_v is not None:
            reason = f"{reason} IRR = {_fmt_num(irr_v, 4)}."
        return ("RECOMMENDED", _BRAND["ok"], reason)

    if npv_v > 0:
        return (
            "CONDITIONAL",
            _BRAND["accent"],
            "Positive NPV but long payback; validate strategic value and risk tolerance.",
        )

    return ("NOT RECOMMENDED", _BRAND["primary"], "Negative NPV under current assumptions.")


def _add_kpi_cards_page(pdf: PdfPages, solution_summary: dict, run_dir: Path, settings_snapshot: dict) -> None:
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.axis("off")

    ax.text(
        0.03,
        0.95,
        "Swiss Post BESS Report - Executive Dashboard",
        fontsize=20,
        fontweight="bold",
        color=_BRAND["secondary"],
        transform=ax.transAxes,
    )
    ax.text(
        0.03,
        0.91,
        f"Run: {run_dir.name}   |   Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        fontsize=10,
        color=_BRAND["neutral"],
        transform=ax.transAxes,
    )
    ax.text(
        0.03,
        0.88,
        f"Optimization mode: {settings_snapshot.get('parameters', {}).get('optimization_mode', 'n/a')}",
        fontsize=10,
        color=_BRAND["neutral"],
        transform=ax.transAxes,
    )

    status, color, reason = _get_recommendation(solution_summary)
    ax.text(
        0.03,
        0.835,
        f"Recommendation: {status}",
        fontsize=11,
        color="white",
        bbox={"boxstyle": "round,pad=0.35", "facecolor": color, "edgecolor": "none"},
        transform=ax.transAxes,
    )
    wrapped_reason = textwrap.wrap(reason, width=52)
    for idx, line in enumerate(wrapped_reason):
        ax.text(0.33, 0.79 - idx * 0.036, line, fontsize=9.3, color=_BRAND["neutral"], transform=ax.transAxes)

    cards = [
        ("Battery Size", f"{_fmt_num(solution_summary.get('battery_capacity_kwh'))} kWh"),
        ("Battery Power", f"{_fmt_num(solution_summary.get('battery_power_capacity_kw'))} kW"),
        ("TAC", f"{_fmt_num(solution_summary.get('objective_total_cost'))} CHF/year"),
        ("NPV", f"{_fmt_num(solution_summary.get('npv'))} CHF"),
        ("IRR", _fmt_num(solution_summary.get("irr"), 4)),
        ("Payback", f"{_fmt_num(solution_summary.get('payback_years'))} years"),
        ("Discounted Payback", f"{_fmt_num(solution_summary.get('discounted_payback_years'))} years"),
        ("No-Battery Cost", f"{_fmt_num(solution_summary.get('no_battery_total_cost'))} CHF/year"),
    ]

    n_cols = 4
    card_w = 0.215
    card_h = 0.16
    x0 = 0.03
    y0 = 0.56
    x_gap = 0.02
    y_gap = 0.06

    for idx, (label, value) in enumerate(cards):
        r = idx // n_cols
        c = idx % n_cols
        x = x0 + c * (card_w + x_gap)
        y = y0 - r * (card_h + y_gap)
        rect = plt.Rectangle(
            (x, y),
            card_w,
            card_h,
            facecolor="#F8FAFD",
            edgecolor=_BRAND["grid"],
            linewidth=1.0,
            transform=ax.transAxes,
        )
        ax.add_patch(rect)
        label_lines = textwrap.wrap(label, width=18)[:2]
        ax.text(
            x + 0.015,
            y + card_h - 0.05,
            "\n".join(label_lines),
            fontsize=10,
            color=_BRAND["neutral"],
            transform=ax.transAxes,
        )
        ax.text(
            x + 0.015,
            y + 0.045,
            _safe_text(value, width=14, max_lines=2),
            fontsize=10.2,
            fontweight="bold",
            color=_BRAND["secondary"],
            transform=ax.transAxes,
        )

    # Cost delta callout
    try:
        baseline = float(solution_summary.get("no_battery_total_cost"))
        tac = float(solution_summary.get("objective_total_cost"))
        delta = baseline - tac
        delta_txt = f"Estimated annual saving vs no battery: {_fmt_num(delta)} CHF/year"
    except Exception:
        delta_txt = "Estimated annual saving vs no battery: n/a"
    ax.text(
        0.03,
        0.08,
        delta_txt,
        fontsize=11,
        color=_BRAND["secondary"],
        fontweight="bold",
        transform=ax.transAxes,
    )

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def generate_pdf_report(run_dir: Path, solution_summary: dict, settings_snapshot: dict, input_dict: dict) -> Path:
    run_dir = Path(run_dir)
    report_path = run_dir / "results_report.pdf"

    # Load optional CSV artifacts if present.
    ts_path = run_dir / "timeseries_results.csv"
    monthly_path = run_dir / "monthly_summary.csv"
    weekly_path = run_dir / "weekly_summary.csv"
    compare_path = run_dir / "baseline_vs_optimized.csv"
    peak_metrics_path = run_dir / "peak_metrics.csv"
    top_peaks_path = run_dir / "top_peak_intervals.csv"
    util_path = run_dir / "battery_utilization_summary.csv"
    fin_path = run_dir / "financial_cashflows.csv"
    sens_path = run_dir / "battery_size_sensitivity.csv"

    df_ts = pd.read_csv(ts_path, parse_dates=["timestamp"]) if ts_path.exists() else pd.DataFrame()
    df_monthly = pd.read_csv(monthly_path) if monthly_path.exists() else pd.DataFrame()
    df_weekly = pd.read_csv(weekly_path) if weekly_path.exists() else pd.DataFrame()
    df_compare = pd.read_csv(compare_path) if compare_path.exists() else pd.DataFrame()
    df_peak_metrics = pd.read_csv(peak_metrics_path) if peak_metrics_path.exists() else pd.DataFrame()
    df_top_peaks = pd.read_csv(top_peaks_path) if top_peaks_path.exists() else pd.DataFrame()
    df_util = pd.read_csv(util_path) if util_path.exists() else pd.DataFrame()
    df_fin = pd.read_csv(fin_path) if fin_path.exists() else pd.DataFrame()
    df_sens = pd.read_csv(sens_path) if sens_path.exists() else pd.DataFrame()

    with PdfPages(report_path) as pdf:
        # 1) Cover + executive summary
        _add_kpi_cards_page(pdf, solution_summary, run_dir, settings_snapshot)
        title_lines = [
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Run folder: {run_dir.name}",
            f"Optimization mode: {settings_snapshot.get('parameters', {}).get('optimization_mode', 'n/a')}",
            f"Timesteps used: {settings_snapshot.get('timesteps_used', 'n/a')}",
            "",
            "Core KPIs",
            f"- Battery Capacity [kWh]: {_fmt_num(solution_summary.get('battery_capacity_kwh'))}",
            f"- Battery Power Capacity [kW]: {_fmt_num(solution_summary.get('battery_power_capacity_kw'))}",
            f"- TAC / Objective [CHF/year]: {_fmt_num(solution_summary.get('objective_total_cost'))}",
            f"- Import Cost [CHF/year]: {_fmt_num(solution_summary.get('import_cost'))}",
            f"- Peak Demand Cost [CHF/year]: {_fmt_num(solution_summary.get('peak_demand_cost'))}",
            f"- No-battery Total Cost [CHF/year]: {_fmt_num(solution_summary.get('no_battery_total_cost'))}",
            f"- NPV [CHF]: {_fmt_num(solution_summary.get('npv'))}",
            f"- IRR [-]: {_fmt_num(solution_summary.get('irr'), digits=4)}",
            f"- Payback [years]: {_fmt_num(solution_summary.get('payback_years'))}",
            f"- Discounted Payback [years]: {_fmt_num(solution_summary.get('discounted_payback_years'))}",
        ]
        _add_text_page(
            pdf,
            "Swiss Post BESS Results Report",
            title_lines,
            subtitle="Automated decision-support report for BESS sizing and dispatch",
        )
        _add_text_page(
            pdf,
            "Executive Insights",
            _build_executive_insights(solution_summary),
            subtitle="One-page interpretation for decision-makers",
        )
        # 2) KPI and baseline comparison tables
        _add_table_page(
            pdf,
            "Key Performance Indicators",
            _build_kpi_table(solution_summary),
            max_rows=20,
        )
        _add_table_page(
            pdf,
            "Baseline vs Optimized Comparison",
            _to_report_table(
                df_compare,
                keep_columns=["Metric", "Baseline", "Optimized", "Optimized - Baseline", "Unit"],
                decimals=2,
            ),
            max_rows=18,
            subtitle="Financial and operational deltas against no-battery baseline",
        )

        # 3) Technical behavior and operations
        _plot_grid_import_timeseries(pdf, df_ts)
        _plot_duration_curve(pdf, df_ts)
        _plot_representative_peak_days(pdf, df_ts)
        _plot_representative_dispatch_weeks(pdf, df_ts)
        _plot_dispatch_and_soc(pdf, df_ts)
        _add_table_page(pdf, "Peak Metrics", _to_report_table(df_peak_metrics, ["Metric", "Before battery", "After battery", "Reduction", "Reduction %", "Unit"]), max_rows=12)
        _add_table_page(
            pdf,
            "Top Peak Intervals",
            _to_report_table(
                df_top_peaks,
                keep_columns=["timestamp", "baseline_grid_import", "optimized_grid_import", "peak_reduction"],
                rename_map={"timestamp": "Timestamp", "baseline_grid_import": "Baseline [kW]", "optimized_grid_import": "Optimized [kW]", "peak_reduction": "Reduction [kW]"},
            ).head(8),
            max_rows=12,
        )
        _add_table_page(pdf, "Battery Utilization", _to_report_table(df_util, ["Metric", "Value", "Unit"]), max_rows=12)

        # 4) Monthly/weekly business impact
        _plot_monthly_savings(pdf, df_monthly)
        _plot_weekly_savings(pdf, df_weekly)
        _add_table_page(
            pdf,
            "Monthly Summary",
            _to_report_table(
                df_monthly,
                keep_columns=["month", "monthly_import_cost_before", "monthly_import_cost_after", "monthly_savings", "monthly_peak_reduction"],
                rename_map={
                    "month": "Month",
                    "monthly_import_cost_before": "Before [CHF]",
                    "monthly_import_cost_after": "After [CHF]",
                    "monthly_savings": "Savings [CHF]",
                    "monthly_peak_reduction": "Peak reduction [kW]",
                },
            ),
            max_rows=14,
        )
        _add_table_page(pdf, "Weekly Highlights", _build_weekly_highlights(df_weekly), max_rows=14)

        # 5) Financial section
        _plot_cashflows(pdf, df_fin)
        _add_table_page(pdf, "Financial Cashflows", _to_report_table(df_fin, ["year", "cashflow", "discounted_cashflow"]).head(12), max_rows=14)

        # 6) Sensitivity section
        _plot_sensitivity(pdf, df_sens)
        _add_table_page(
            pdf,
            "Battery Size Sensitivity Table",
            _to_report_table(
                df_sens,
                keep_columns=["battery_size_kwh", "objective_total_cost", "npv", "payback_years", "status"],
                rename_map={
                    "battery_size_kwh": "Battery size [kWh]",
                    "objective_total_cost": "TAC [CHF/year]",
                    "npv": "NPV [CHF]",
                    "payback_years": "Payback [years]",
                    "status": "Status",
                },
            ),
            max_rows=14,
        )

        # 7) Settings and assumptions appendix
        settings_table = _build_settings_table(settings_snapshot, input_dict)
        _add_table_page(
            pdf,
            "Run Settings and Model Parameters",
            settings_table,
            max_rows=16,
            subtitle="Configuration snapshot used for this optimization run",
        )
        assumptions_lines = [
            "Key modeling assumptions and settings",
            f"- Surplus handling: {input_dict.get('parameters', {}).get('surplus_handling', 'n/a')}",
            f"- Peak tariff granularity: {input_dict.get('parameters', {}).get('peak_shaving_frequency', 'n/a')}",
            f"- Peak tariff factor: {_fmt_num(input_dict.get('parameters', {}).get('peak_shaving_cost_factor', 'n/a'))}",
            f"- Energy CAPEX [CHF/kWh]: {_fmt_num(input_dict.get('parameters', {}).get('Battery_energy_invest_cost', 'n/a'))}",
            f"- Power CAPEX [CHF/kW]: {_fmt_num(input_dict.get('parameters', {}).get('Battery_power_invest_cost', 'n/a'))}",
            f"- Interest rate [-]: {_fmt_num(input_dict.get('parameters', {}).get('interest_rate', 'n/a'), digits=4)}",
            f"- Lifetime [years]: {_fmt_num(input_dict.get('parameters', {}).get('lifetime', 'n/a'), digits=0)}",
            "",
            "Notes",
            "- TAC denotes total annualized cost and is the optimization objective.",
            "- Sensitivity includes infeasible points when present.",
        ]
        _add_text_page(pdf, "Assumptions and Notes", assumptions_lines)

    return report_path

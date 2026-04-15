# CS1-Swiss-Post

Battery sizing and dispatch optimization for a Swiss Post case study.

This repository builds 15-minute site time series, optimizes BESS sizing and operation, and exports technical and financial results for decision support.

## Purpose and Decisions Supported

Use this model when you want to answer:

- Is a BESS economically attractive at this site under current assumptions?
- What battery energy size (`kWh`) and power size (`kW`) are cost-optimal?
- How much value comes from reduced import cost and peak demand cost?
- What is the expected business case (NPV, payback, IRR)?

Current primary value streams modeled:

- Peak shaving (yearly or monthly demand-charge style penalty)
- PV self-consumption support (through dispatch and curtailment handling)

Not modeled as objective value streams right now:

- Grid export remuneration
- Electricity arbitrage
- Ancillary services

## Quick Start

1. Create and activate a Python environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Set parameters in `config.py`.
4. Run:

```bash
python main.py
```

## Run Flow

`main.py` orchestrates:

1. Load/build `03-PROCESSED-DATA/input_dict.json`
2. Optional horizon cap (`max_timesteps`)
3. Build model (`optimization.py`)
4. Solve and summarize solution
5. Compute no-battery baseline
6. Compute financial metrics (`results_processing.py`)
7. Export run artifacts to `02-MODEL-RESULTS/<timestamp>_<mode>_<nsteps>steps/`

## Input Data Contract

### Required Excel structure

Data should follow the template workbook format:

- Trafo and PV:
  - timestamp column `Zeit`
  - 10-minute resolution
  - power column header containing `-avg[W]`
- EV charging profile (LKW):
  - time column in `HH:MM:SS`
  - one full weekday profile in 15-minute resolution
  - Saturday/Sunday assumed zero
- Delivery profile (Zustellung):
  - full-day template (`00:00` to `23:45`)
  - first column time values
  - columns for Monday to Saturday in `kW`
  - Sunday assumed zero
  - summer reduction applied in preprocessing

### Data loading behavior

If `load_existing_input_dict=False`, data is selected interactively via dialogs in `data_preprocessing.py`.

## Modeling Definitions and Sign Conventions

Core power-balance constraint per timestep:

`grid_flow + pv_out_flow + battery_out_flow - battery_in_flow = total_demand`

Implemented variable signs:

- `grid_flow >= 0` (import only)
- `pv_out_flow >= 0`
- `battery_in_flow >= 0` (charging power)
- `battery_out_flow >= 0` (discharging power)
- `battery_level >= 0`

Important interpretation:

- `total_demand` is an input time series assembled in preprocessing.
- In reporting notebooks, this is often interpreted as **net load**.
- A derived **actual load** (including battery charging) can be computed as:
  - `actual_load = grid_flow + pv_out_flow + battery_out_flow`
  - equivalent to `actual_load = total_demand + battery_in_flow`

## Optimization Formulation (Current)

Solver: OR-Tools CBC (`ortools.linear_solver.pywraplp`)

Decision variables:

- battery capacity (`kWh`)
- battery power capacity (`kW`)
- battery charge/discharge trajectories
- grid import trajectory
- PV use and curtailment
- battery state of charge

Objective:

`Total_Cost = Annualized_Battery_Cost + O&M + Import_Cost + Peak_Demand_Cost`

Where:

- `Import_Cost = sum(grid_flow[t] * electricity_price[t] * 0.25)`
- `Annualized_Battery_Cost = CRF * (energy_capex * Battery_capacity + power_capex * Battery_power_capacity)`
- `Peak_Demand_Cost` is yearly or monthly based on config

Modes:

- `optimization_mode="milp"`: binary exclusivity for charge/discharge
- `optimization_mode="lp"`: faster, but may allow simultaneous charge/discharge

## Configuration Guide

Main parameters in `config.py`:

- Scenario and runtime:
  - `load_existing_input_dict`
  - `max_timesteps`
  - `optimization_mode`
  - `surplus_handling`
- Battery physics and bounds:
  - `Battery_max_inflow`
  - `Battery_max_outflow`
  - `Battery_max_capacity`
  - `battery_max_c_rate`
  - `battery_min_soc_fraction`
  - `eta_charge`, `eta_discharge`, `eta_self_discharge`
- Economics:
  - `invest_cost_energy`, `invest_cost_power`
  - `om_fixed_chf_per_year`, `om_energy_chf_per_kwh_year`, `om_power_chf_per_kw_year`
  - `interest_rate`, `lifetime`
  - `peak_shaving_cost_factor`, `peak_shaving_frequency`
  - replacement assumptions for postprocessing

## Outputs and How to Interpret Them

Each run writes to:

- `02-MODEL-RESULTS/<timestamp>_<mode>_<nsteps>steps/`

Common files:

- `results_summary.json`: headline KPIs
- `timeseries_results.csv`: timestep results
- `baseline_vs_optimized.csv`: side-by-side cost comparison
- `financial_cashflows.csv`: yearly cashflow and discounted cashflow
- `monthly_summary.csv`: monthly import cost and peak reductions
- `weekly_summary.csv`: weekly energy/cost/peak reductions
- `peak_metrics.csv`, `top_peak_intervals.csv`
- `battery_utilization_summary.csv`
- `battery_size_sensitivity.csv`: battery size sweep results (when enabled)

KPI interpretation notes:

- `objective_total_cost`: annual optimization objective (includes annualized battery cost)
- `no_battery_total_cost`: annual baseline cost without battery
- `annual_savings`: operating savings used in financial postprocessing
- `npv`, `irr`, `payback_years`: calculated in postprocessing, not embedded as MILP objective

Sensitivity analysis notes:

- Activate with `run_battery_size_sensitivity=True` in `config.py`.
- If `battery_sensitivity_sizes_kwh` is empty, default size points are generated automatically around the optimized battery size.
- Reporting in `results_sheet.ipynb` uses TAC (`objective_total_cost`) as primary sensitivity axis, with NPV on the secondary axis.
- Infeasible size points are retained and highlighted separately in the sensitivity plot.

## Validation Checklist (Recommended Every Run)

Before trusting a run, check:

- Solve status is optimal.
- `timeseries_results.csv` has expected timestep count.
- No obvious NaNs in key columns (`grid_flow`, `battery_soc`, `total_load`).
- Battery dispatch is plausible (SOC within bounds, charge/discharge behavior reasonable).
- Peak metrics trend makes sense when peak charges are enabled.
- Baseline vs optimized comparison direction is economically reasonable.

For reporting:

- Confirm whether charts use **net load** or **actual load** definitions.
- Use consistent terminology in slides and report tables.

## Limitations and Decision Boundaries

Current limitations:

- Import-only grid model (`grid_flow >= 0`), no export remuneration.
- Electricity price is currently constant when rebuilding input dictionary (`0.30 CHF/kWh` in `main.py`).
- LP mode may produce non-physical simultaneous charge/discharge.
- Preprocessing is GUI-driven and not batch-friendly for large multi-site automation.

Implication for decisions:

- Good for comparative BESS screening and sensitivity at site level.
- Not yet suitable for export/arbitrage market strategy assessment.

## Practical Notes

- `load_existing_input_dict=True` requires `03-PROCESSED-DATA/input_dict.json`.
- If missing, run once with `load_existing_input_dict=False`.
- `max_timesteps` is useful for fast debugging runs.
- `data_preprocessing.py` currently rewrites `config.py` from Excel sheet `config` at import time.
- Repository currently contains two config workflows (`config.py` and `config_new.py` with `CONFIG_INPUTS`), while `main.py` uses `config.py`.

## Key Files

- `main.py`: run orchestration and artifact creation
- `optimization.py`: model variables, constraints, objective, solver call
- `data_preprocessing.py`: Excel ingestion and profile construction
- `results_processing.py`: baseline, financial metrics, and exports
- `results_sheet.ipynb`: reporting notebook

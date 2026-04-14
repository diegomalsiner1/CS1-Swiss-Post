# CS1 Swiss Post

Battery sizing and dispatch optimization for a Swiss Post case study.

This repository builds a 15-minute site-level power dataset from Excel inputs, solves a linear optimization model for battery operation and capacity, and exports technical and financial KPIs.

## Repository Overview

Main production files:

- `main.py`: end-to-end orchestration (load/build input data, solve optimization, export artifacts)
- `data_preprocessing.py`: Excel ingestion, cleaning, 10-minute to 15-minute conversion, EV profile generation
- `optimization.py`: OR-Tools LP/MILP formulation and solver call
- `results_processing.py`: KPI tables, financial analysis, and CSV/JSON exports
- `config.py`: runtime parameters currently consumed by `main.py`

Supporting files:

- `config_new.py`: alternate config loader from Excel `CONFIG_INPUTS` sheet (not used by `main.py` by default)
- `first_test.py`, `Test.py`: exploratory/test scripts
- `prepare_input_excel.ipynb`: creates a template workbook with a `CONFIG_INPUTS` sheet
- `visualize_data_processed.ipynb`: quick checks/plots for `03-PROCESSED-DATA/data_processed.csv`
- `results_sheet.ipynb`: report-style visualization for the latest optimization run

Data/result folders:

- `01-INPUT-DATA/`: raw Excel inputs and templates
- `02-MODEL-RESULTS/`: timestamped run outputs
- `03-PROCESSED-DATA/`: generated intermediate artifacts (`data_processed.csv`, `input_dict.json`)

## Environment Setup

1. Create and activate a Python environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Input Data Requirements

The main workflow expects a single Excel workbook selected through a file dialog.

### Required sheet groups

1. Transformer sheets (select one or more):
- Must contain a time column named `Zeit`
- Must contain at least one power column with `-avg[W]` in the header
- Expected in 10-minute resolution

2. PV sheet (select one):
- Same `Zeit` + `-avg[W]` expectation as transformer sheets

3. EV charging sheet (`LKW`, select one):
- Expected as a 15-minute day profile with 96 valid rows after parsing
- Parsed from Excel formulas/values by `data_preprocessing.py`

4. Delivery profile sheet (`Zustellung`, select one):
- First column must be time (`HH:MM:SS` style)
- Columns for weekdays `Montag` to `Samstag`
- Sunday is assumed zero

### Important preprocessing behavior

- Transformer and PV series are aligned and converted from 10-minute to 15-minute with energy-conserving conversion.
- LKW profile is expanded over weekdays for the full year.
- Zustellung profile is expanded over the full year (Mon-Sat), with April-September reduced to 60%.
- Output is written to `03-PROCESSED-DATA/data_processed.csv`.

## Configuration

`main.py` imports `config.py`.

Key parameters in `config.py` include:

- `load_existing_input_dict`
- `max_timesteps`
- `optimization_mode` (`lp` or `milp`)
- `PV_max_capacity`
- `Battery_max_inflow`, `Battery_max_outflow`, `Battery_max_capacity`
- `eta_charge`, `eta_discharge`, `eta_self_discharge`
- `invest_cost`, `operation_and_maintenance`
- `interest_rate`, `lifetime`
- `peak_shaving_cost_factor`
- `peak_shaving_frequency` (or `peak_shaving_granularity` key in input dictionary)

Notes:

- `load_existing_input_dict = True` requires `03-PROCESSED-DATA/input_dict.json` to exist.
- If it does not exist, run once with `load_existing_input_dict = False`.
- `max_timesteps` can be used to shorten runs for debugging.

## How To Run

```bash
python main.py
```

When `load_existing_input_dict = False`, the script will open dialogs to:

1. Select the Excel file
2. Select transformer sheets
3. Select PV sheet
4. Select EV charging sheet
5. Select delivery profile sheet

## Model Summary

Solver backend:

- OR-Tools CBC (`ortools.linear_solver.pywraplp`)

Decision variables per timestep include:

- Battery charge/discharge power
- Battery state of charge
- Grid import
- PV dispatch

Global decision variable:

- Battery capacity

Objective minimized:

- `annualized_battery_cost + operation_and_maintenance + import_cost + peak_demand_cost`

Modes:

- `lp`: continuous model (faster; can allow simultaneous charge/discharge)
- `milp`: adds binary exclusivity constraints for battery in/out flow

## Baseline and Financial Outputs

A no-battery baseline is computed on the same demand, PV availability, and tariffs.

Financial post-processing computes and exports:

- annual savings
- NPV
- simple payback
- discounted payback

## Output Files

Each run creates a folder under:

- `02-MODEL-RESULTS/<timestamp>_<mode>_<nsteps>steps/`

Typical artifacts:

- `settings_snapshot.json`
- `results_summary.json`
- `timeseries_results.csv`
- `financial_cashflows.csv`
- `baseline_vs_optimized.csv`
- `peak_metrics.csv`
- `top_peak_intervals.csv`
- `monthly_summary.csv`
- `battery_utilization_summary.csv`
- `battery_size_sensitivity.csv` (if enabled)

Intermediate artifacts:

- `03-PROCESSED-DATA/data_processed.csv`
- `03-PROCESSED-DATA/input_dict.json`

## Known Caveats

- `data_preprocessing.py` currently rewrites `config.py` at import time from an Excel sheet named `config`.
- `main.py` uses `config.py`, while `prepare_input_excel.ipynb` and `config_new.py` are built around a `CONFIG_INPUTS` sheet. This means there are two config workflows in the repository.
- Grid export is not modeled (grid variable is non-negative import only).
- Electricity price is currently set as a constant `0.30 CHF/kWh` when building fresh input data.

## Suggested Workflow

1. Place and format source Excel data in `01-INPUT-DATA/`.
2. Set parameters in `config.py`.
3. Run `python main.py`.
4. Inspect outputs in the newest folder under `02-MODEL-RESULTS/`.
5. Use notebooks for reporting and visualization.

# CS1 Swiss Post

Battery sizing and dispatch optimization for a Swiss Post case study.

This repository builds a 15-minute site-level power dataset from Excel inputs, solves a linear optimization model for battery operation and capacity, and exports technical and financial KPIs.

## What This Model Does

The optimization decides:

- Battery capacity (`kWh`)
- Battery power rating (`kW`)
- Battery charging/discharging trajectory (`kW` per timestep)
- Grid import (`kW` per timestep)
- PV usage (`kW` per timestep)

Objective:

- Minimize total annualized cost
- `Total_Cost = Annualized_Battery_Cost + Battery_O&M + Import_Cost + Peak_Demand_Cost`

The model runs on a fixed 15-minute timestep (`0.25 h`).

## Optimization Modes

Configured in `config.py` with `optimization_mode`:

- `"milp"`: adds binary variables to prevent simultaneous charge/discharge
- `"lp"`: removes binary variables (faster solve, but simultaneous charge/discharge can occur)

In LP mode, these binary variables and related constraints are disabled:

- `Binary_battery_in_flow[t]`
- `Binary_battery_out_flow[t]`
- Mutual exclusivity constraints linking battery flows to binaries

Solver backend: OR-Tools CBC (`ortools.linear_solver.pywraplp`).

## Data Flow

End-to-end pipeline in `main.py`:

1. Load or generate input dictionary (`03-PROCESSED-DATA/input_dict.json`)
2. Optionally cap horizon with `max_timesteps`
3. Build optimization model (`optimization.py`)
4. Solve model and summarize KPIs
5. Compute no-battery baseline
6. Export run artifacts in timestamped folder under `02-MODEL-RESULTS/`

## Input Data Sources

When `load_existing_input_dict = False`, raw data is read from an Excel file selected via GUI dialog (`tkinter`):

- Transformer/grid data via selected sheets
- PV data via selected PV sheet
- EV profiles:
	- `Zubau_LKW_Aug2025` (weekday profile repeated over year)
	- `Zubau_Zustellung_Oct2026` (weekday-dependent profile, Sunday=0, summer reduction)

Preprocessing behavior in `data_preprocessing.py`:

- Trafo data cleaned and aligned
- 10-minute series converted to 15-minute (energy-conserving conversion)
- EV profiles expanded to full-year 15-minute series
- Merged output written to `03-PROCESSED-DATA/data_processed.csv`

Input dictionary structure (core keys):

- `parameters`
- `timestamps`
- `total_demand`
- `PV_capacity_factor`
- `electricity_price`

## Configuration

Main knobs in `config.py`:

- `load_existing_input_dict`: use cached processed input (`True/False`)
- `max_timesteps`: horizon cap (`None` or positive int)
- `optimization_mode`: `"lp"` or `"milp"`
- `surplus_handling`: `"curtail"` or `"must_absorb"`
- `PV_max_capacity` (`kW`)
- `Battery_max_inflow` (`kW`)
- `Battery_max_outflow` (`kW`)
- `Battery_max_capacity` (`kWh`) upper bound
- `battery_max_c_rate` (`1/h`) power-to-energy coupling for charge and discharge
- `battery_min_soc_fraction` (`-`) minimum state of charge reserve
- `eta_charge`, `eta_discharge`, `eta_self_discharge`
- `invest_cost_energy` (`CHF/kWh`) energy-capacity CAPEX
- `invest_cost_power` (`CHF/kW`) power-rating CAPEX
- `battery_cycle_life` (equivalent full cycles before replacement)
- `battery_calendar_life_years` optional calendar-life cap
- `battery_replacement_cost_fraction` (`-`) replacement CAPEX as share of initial CAPEX
- `operation_and_maintenance` (`CHF`)
- `interest_rate`, `lifetime` (for annualization)

Backward compatibility:

- older cached input dictionaries/results may still contain `Battery_invest_cost` or `battery_replacement_year`; the code still reads them as legacy fallbacks

## Model Formulation (Implemented)

Decision variables per timestep `t`:

- `Battery_in_flow[t] >= 0`
- `Battery_out_flow[t] >= 0`
- `Grid_flow[t] >= 0`
- `PV_out_flow[t] >= 0`
- `Spill_flow[t] >= 0`
- `Battery_level[t] >= 0`

Global variable:

- `Battery_capacity` with upper bound `Battery_max_capacity`
- `Battery_power_capacity` with upper bound `max(Battery_max_inflow, Battery_max_outflow)`

Core constraints:

- Power balance:
	- `Grid + PV + Battery_out - Battery_in = demand`
- PV allocation:
	- `PV_out + Spill = PV_capacity_factor[t] * PV_max_capacity`
	- if `surplus_handling = "must_absorb"`, then `Spill = 0`
- Battery state dynamics:
	- `SOC[t+1] = SOC[t]*(1-self_discharge) + 0.25*eta_charge*Battery_in - 0.25*(1/eta_discharge)*Battery_out`
- SOC bounded by capacity:
	- `Battery_level[t] <= Battery_capacity`
- Minimum SOC reserve:
	- `Battery_level[t] >= battery_min_soc_fraction * Battery_capacity`
- Optional C-rate coupling:
	- `Battery_in_flow[t] <= battery_max_c_rate * Battery_capacity`
	- `Battery_out_flow[t] <= battery_max_c_rate * Battery_capacity`
- Power-rating coupling:
	- `Battery_in_flow[t] <= Battery_power_capacity`
	- `Battery_out_flow[t] <= Battery_power_capacity`
	- if `battery_max_c_rate` is set: `Battery_power_capacity <= battery_max_c_rate * Battery_capacity`
	- otherwise a fallback linear linkage keeps power rating at zero when energy capacity is zero
- Cyclic SOC boundary:
	- `SOC[0] = 0.5*Battery_capacity`
	- `SOC[last] = 0.5*Battery_capacity`

MILP-only constraints:

- `Binary_in[t] + Binary_out[t] <= 1`
- flow-binary linking constraints

Objective terms:

- `Import_Cost = sum(Grid_flow[t] * price[t] * 0.25)`
- `Annualized_Battery_Cost = CRF(interest_rate, lifetime) * (invest_cost_energy * Battery_capacity + invest_cost_power * Battery_power_capacity)`
- `Peak_Demand_Cost = peak_shaving_cost_factor * yearly_peak` or monthly sum of peaks
- `Total_Cost = Annualized_Battery_Cost + operation_and_maintenance + Import_Cost + Peak_Demand_Cost`

## Baseline and Financial Metrics

After optimization, the project computes:

- No-battery baseline import cost (`no_battery_import_cost`)
- No-battery baseline peak-demand cost (`no_battery_peak_demand_cost`)
- No-battery total cost (`no_battery_total_cost`)
- NPV, IRR, and payback from annual operating savings, with upfront CAPEX and optional replacement cashflow treated separately (`results_processing.py`)
- CAPEX split into energy and power components in the exported comparison tables
- Replacement timing is estimated from annual equivalent full cycles and `battery_cycle_life`, optionally capped by `battery_calendar_life_years`

Financial cashflow table is exported as:

- `financial_cashflows.csv`

## How To Run

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

- `settings_snapshot.json` (run settings and parameters)
- `results_summary.json` (core KPIs)
- `timeseries_results.csv` (aligned timestep results)
- `financial_cashflows.csv` (yearly financial cashflows)
- `battery_size_sensitivity.csv` (if sensitivity is enabled)

Processed input artifacts:

- `03-PROCESSED-DATA/input_dict.json`
- `03-PROCESSED-DATA/data_processed.csv`

## Practical Notes

- `load_existing_input_dict=True` expects `03-PROCESSED-DATA/input_dict.json` to exist.
- If not present, set `load_existing_input_dict=False` for one run to generate it.
- `max_timesteps` can speed up debugging and test runs.
- Preprocessing uses GUI sheet/file dialogs; this is not headless-server friendly.
- CBC output is enabled in solve phase, so terminal logs can be verbose.
- `data_preprocessing.py` currently rewrites `config.py` at import time from an Excel sheet named `config`.
- The repository contains two config workflows (`config.py` and `config_new.py` / `CONFIG_INPUTS`), while `main.py` currently uses `config.py`.
- Sensitivity analysis can be enabled via `run_battery_size_sensitivity=True` in `config.py`.
- If `battery_sensitivity_sizes_kwh` is left empty, the model now generates a default size grid around the optimized battery size automatically.
- In `results_sheet.ipynb`, sensitivity plots use TAC (`objective_total_cost`) with NPV on a second axis; infeasible points are shown separately.

## Current Scope and Limitations

- Grid flow is modeled as non-negative import only (no export variable).
- Curtailment/spillage can be allowed via `surplus_handling = "curtail"` without any export revenue.
- Electricity price is currently set as constant `0.30 CHF/kWh` during input generation.
- LP mode may allow physically unrealistic simultaneous charging/discharging.

## Key Files

- `main.py`: pipeline orchestration and artifact management
- `optimization.py`: model variables, constraints, objective, solve
- `data_preprocessing.py`: Excel ingestion, cleaning, profile generation
- `results_processing.py`: KPI financial post-processing and exports
- `config.py`: scenario and model parameters

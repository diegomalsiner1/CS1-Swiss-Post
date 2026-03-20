# CS1-Swiss-Post

Battery sizing and dispatch optimization for a Swiss Post case study.

This project builds a 15-minute load/PV dataset, solves a linear optimization model for battery operation and sizing, and exports technical and financial KPIs.

## What This Model Does

The optimization decides:

- Battery capacity (`kWh`)
- Battery charging/discharging trajectory (`kW` per timestep)
- Grid import (`kW` per timestep)
- PV usage (`kW` per timestep)

Objective:

- Minimize total annualized cost
- `Total_Cost = Annualized_Battery_Cost + OPEX`
- `OPEX = Fixed_O&M + Import_Cost`

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
- `PV_max_capacity` (`kW`)
- `Battery_max_inflow` (`kW`)
- `Battery_max_outflow` (`kW`)
- `Battery_max_capacity` (`kWh`) upper bound
- `eta_charge`, `eta_discharge`, `eta_self_discharge`
- `invest_cost` (`CHF/kWh`)
- `operation_and_maintenance` (`CHF`)
- `interest_rate`, `lifetime` (for annualization)

## Model Formulation (Implemented)

Decision variables per timestep `t`:

- `Battery_in_flow[t] >= 0`
- `Battery_out_flow[t] >= 0`
- `Grid_flow[t] >= 0`
- `PV_out_flow[t] >= 0`
- `Battery_level[t] >= 0`

Global variable:

- `Battery_capacity` with upper bound `Battery_max_capacity`

Core constraints:

- Power balance:
	- `Grid + PV + Battery_out - Battery_in = demand`
- PV limit:
	- `PV_out <= PV_capacity_factor[t] * PV_max_capacity`
- Battery state dynamics:
	- `SOC[t+1] = SOC[t]*(1-self_discharge) + 0.25*eta_charge*Battery_in - 0.25*(1/eta_discharge)*Battery_out`
- SOC bounded by capacity:
	- `Battery_level[t] <= Battery_capacity`
- Cyclic SOC boundary:
	- `SOC[0] = 0.5*Battery_capacity`
	- `SOC[last] = 0.5*Battery_capacity`

MILP-only constraints:

- `Binary_in[t] + Binary_out[t] <= 1`
- flow-binary linking constraints

Objective terms:

- `Import_Cost = sum(Grid_flow[t] * price[t] * 0.25)`
- `Annualized_Battery_Cost = CRF(interest_rate, lifetime) * invest_cost * Battery_capacity`
- `Total_Cost = Annualized_Battery_Cost + operation_and_maintenance + Import_Cost`

## Baseline and Financial Metrics

After optimization, the project computes:

- No-battery baseline import cost (`no_battery_import_cost`)
- No-battery total cost (`no_battery_total_cost`)
- NPV and simple payback from annual savings (`results_processing.py`)

Financial cashflow table is exported as:

- `financial_cashflows.csv`

## How To Run

1. Create and activate a Python environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Set desired parameters in `config.py`.
4. Run:

```bash
python main.py
```

## Output Structure

Each run creates a folder:

- `02-MODEL-RESULTS/<timestamp>_<mode>_<nsteps>steps/`

Typical contents:

- `settings_snapshot.json` (run settings and parameters)
- `results_summary.json` (core KPIs)
- `timeseries_results.csv` (aligned timestep results)
- `financial_cashflows.csv` (yearly financial cashflows)

Processed input artifacts:

- `03-PROCESSED-DATA/input_dict.json`
- `03-PROCESSED-DATA/data_processed.csv`

## Practical Notes

- `load_existing_input_dict=True` expects `03-PROCESSED-DATA/input_dict.json` to exist.
- If not present, set `load_existing_input_dict=False` for one run to generate it.
- `max_timesteps` can speed up debugging and test runs.
- Preprocessing uses GUI sheet/file dialogs; this is not headless-server friendly.
- CBC output is enabled in solve phase, so terminal logs can be verbose.

## Current Scope and Limitations

- Grid flow is modeled as non-negative import only (no export variable).
- Electricity price is currently set as constant `0.30 CHF/kWh` during input generation.
- LP mode may allow physically unrealistic simultaneous charging/discharging.

## Key Files

- `main.py`: pipeline orchestration and artifact management
- `optimization.py`: model variables, constraints, objective, solve
- `data_preprocessing.py`: Excel ingestion, cleaning, profile generation
- `results_processing.py`: KPI financial post-processing and exports
- `config.py`: scenario and model parameters
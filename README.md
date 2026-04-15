# CS1-Swiss-Post
Instructions on file import:

File formatting:

1. Data must be entered into an Excel file that follows the structure of "input_template.xlsx".
2. For the Trafo and PV data:
    2.1 Timestamps must be stored under column "Zeit".
    2.2 Timestamps must be stored in 10 minute intervals.
    2.3 Column title must include "-avg[W]" for power values.
3. Amount of trafo input sheets is arbitrary.
4. For the EV charging profile (LKW):
    4.1 Data must contain a time column in HH:MM:SS format.
    4.2 Data must contain a column with the exact name "Total kW".
    4.3 time column must be in 15-minute resolution covering one full weekday example.
    4.5 Saturday and Sunday are assumed to be 0 load.
5. For the delivery data (Zustellung):
    5.1 Data must be provided as a full-day template (00:00–23:45).
    5.2 First column must contain time values in HH:MM:SS format.
    5.3 One column must exist for each weekday from monday to saturday containing power in kW.
    5.5 Sunday is not included and is assumed to be 0.
    5.6 Data represents winter baseline conditions (summer adjustment is applied in processing).


Configuration:

1. Configurable variables are in a separate worksheet in the input excel file under "config" (as in "input_template.xlsx"). 
2. Only ever change entries in the "values" column. 
3. Expected value types are mentioned in "comments" column in the respective row.


Data Selection:

1. Upon running "main.py", a file dialogue window will be
opened. Select the Excel file with relevant data. The "input_template.xlsx" may be selected to see an example run.
2. A window will open after a short time prompting the user
to select sheets including transformer data. Select all
relevant sheets and click "Apply & Exit".
3. A subsequent window will open prompting the user to select
PV data if available. Select all relevant sheets and click 
"Apply & Exit".
4. The same pattern goes for the PV charging sheet as well as the delivery profile sheet.


Battery sizing and dispatch optimization for a Swiss Post case study.

This project builds a 15-minute load/PV dataset, solves a linear optimization model for battery operation and sizing, and exports technical and financial KPIs.

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
- Curtailment/spillage can be allowed via `surplus_handling = "curtail"` without any export revenue.
- Electricity price is currently set as constant `0.30 CHF/kWh` during input generation.
- LP mode may allow physically unrealistic simultaneous charging/discharging.

## Key Files

- `main.py`: pipeline orchestration and artifact management
- `optimization.py`: model variables, constraints, objective, solve
- `data_preprocessing.py`: Excel ingestion, cleaning, profile generation
- `results_processing.py`: KPI financial post-processing and exports
- `config.py`: scenario and model parameters

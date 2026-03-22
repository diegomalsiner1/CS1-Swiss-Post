import json
from pathlib import Path

## Will be used to define the input parameters

use_case = "Peak_Shaving" ## can be switched between "Peak_Shaving"
interest_rate = 0.06 # [-]
lifetime = 20 # [years]
year =  2025 # 

load_existing_input_dict = True # [True or False]
max_timesteps = 40000 # [None or int] Limit optimization horizon for faster/debug runs
optimization_mode = "lp" # ["milp" or "lp"]
run_battery_size_sensitivity = True # [True or False]
battery_sensitivity_sizes_kwh = [0, 250, 500, 750, 1000, 1250, 1500, 1750, 2000]

PV_max_capacity = 10000 # [kW]
Battery_max_inflow = 1000 # [kW]
Battery_max_outflow = 1000 # [kW]
Battery_max_capacity = 100000 # [kWh]
eta_charge = 0.9 # [-]
eta_discharge = 0.95 # [-]
eta_self_discharge = 0.0 # [-]
invest_cost = 450 # [CHF/kWh]
operation_and_maintenance = 10000 * (
    max_timesteps / (24 * 4 * 365) if max_timesteps is not None else 1) 
battery_degrading = 0.01 # [% per year]

def load_config(path: str = "config.json") -> dict:
	p = Path(path)
	if not p.exists():
		raise FileNotFoundError(f"config.json not found under: {p.resolve()}")

	with p.open("r", encoding="utf-8") as f:
		cfg = json.load(f)

	return cfg


def print_config(cfg: dict) -> None:
	def g(d, *keys, default=None):
		for k in keys:
			if not isinstance(d, dict) or k not in d:
				return default
			d = d[k]
		return d

	print("\n===== CONFIG =====")
	print(f"use_case: {g(cfg, 'run', 'use_case')}")
	print(f"timestep_minutes: {g(cfg, 'run', 'timestep_minutes')}")
	print(f"horizon_days: {g(cfg, 'run', 'horizon_days')}")
	print(f"solver: {g(cfg, 'run', 'solver')}")

	print(f"discount_rate: {g(cfg, 'economics', 'discount_rate')}")
	print(f"lifetime_years: {g(cfg, 'economics', 'lifetime_years')}")
	print(f"capex_chf_per_kwh: {g(cfg, 'economics', 'capex_chf_per_kwh')}")

	print(f"battery.max_capacity_kwh: {g(cfg, 'battery', 'max_capacity_kwh')}")
	print(f"battery.max_inflow_kw: {g(cfg, 'battery', 'max_inflow_kw')}")
	print(f"battery.max_outflow_kw: {g(cfg, 'battery', 'max_outflow_kw')}")
	print(f"battery.eta_charge: {g(cfg, 'battery', 'eta_charge')}")
	print(f"battery.eta_discharge: {g(cfg, 'battery', 'eta_discharge')}")
	print(f"battery.eta_self_discharge_per_step: {g(cfg, 'battery', 'eta_self_discharge_per_step')}")
	print(f"battery.max_c_rate: {g(cfg, 'battery', 'max_c_rate')}")
	print(f"battery.dod_max: {g(cfg, 'battery', 'dod_max')}")

	print(f"excel_path: {g(cfg, 'inputs', 'excel_path')}")
	print(f"load_sheet: {g(cfg, 'inputs', 'load_sheet')}")
	print(f"pv_sheet: {g(cfg, 'inputs', 'pv_sheet')}")
	print(f"tariff_sheet: {g(cfg, 'inputs', 'tariff_sheet')}")
	print("==================\n")

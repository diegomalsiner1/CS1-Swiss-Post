import json
from pathlib import Path


def load_config(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"config.json nicht gefunden unter: {p.resolve()}")

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
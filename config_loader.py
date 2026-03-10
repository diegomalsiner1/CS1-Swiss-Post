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

    print(f"PV_max_capacity: {g(cfg, 'parameters', 'PV_max_capacity')}")
    print(f"Battery_max_capacity: {g(cfg, 'parameters', 'Battery_max_capacity')}")
    print(f"Battery_max_inflow: {g(cfg, 'parameters', 'Battery_max_inflow')}")
    print(f"Battery_max_outflow: {g(cfg, 'parameters', 'Battery_max_outflow')}")
    print(f"Battery_eta_charge: {g(cfg, 'parameters', 'Battery_eta_charge')}")
    print(f"Battery_eta_discharge: {g(cfg, 'parameters', 'Battery_eta_discharge')}")
    print(f"Battery_eta_self_discharge: {g(cfg, 'parameters', 'Battery_eta_self_discharge')}")
    print(f"Battery_max_c_rate: {g(cfg, 'parameters', 'Battery_max_c_rate')}")
    print(f"Battery_dod_max: {g(cfg, 'parameters', 'Battery_dod_max')}")
    print(f"Battery_invest_cost: {g(cfg, 'parameters', 'Battery_invest_cost')}")
    print(f"operation_and_maintenance: {g(cfg, 'parameters', 'operation_and_maintenance')}")
    print(f"interest_rate: {g(cfg, 'parameters', 'interest_rate')}")
    print(f"lifetime: {g(cfg, 'parameters', 'lifetime')}")
    print(f"battery_degrading: {g(cfg, 'parameters', 'battery_degrading')}")

    print(f"excel_path: {g(cfg, 'inputs', 'excel_path')}")
    print(f"load_sheet: {g(cfg, 'inputs', 'load_sheet')}")
    print(f"pv_sheet: {g(cfg, 'inputs', 'pv_sheet')}")
    print(f"tariff_sheet: {g(cfg, 'inputs', 'tariff_sheet')}")
    print("==================\n")
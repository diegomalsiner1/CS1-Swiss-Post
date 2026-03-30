import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import tkinter as tk
import re
from tkinter import filedialog
from datetime import datetime, time
from openpyxl import load_workbook

## Will be used to process the input data before feeding the data into the optimization framework

_selected_file_path = None

def get_input_file_path():
    """
    Lazily open file picker once and cache the selected Excel path.
    """
    global _selected_file_path
    if _selected_file_path is None:
        _selected_file_path = filedialog.askopenfilename(
            initialdir="01-INPUT-DATA",
            title="Select Input Data",
            filetypes=[("Excel files", ".xlsx .xls")]
        )
        if not _selected_file_path:
            raise FileNotFoundError("No input data file selected.")
    return _selected_file_path


# ==========================================================
# config reading and generation
# ==========================================================
import math

def write_config_py(config, filename="config.py"):
    with open(filename, "w", encoding="utf-8") as f:
        f.write("# Auto-generated from Excel - DO NOT EDIT MANUALLY\n\n")

        for key, value in config.items():

            # handle NaN (float)
            if isinstance(value, float) and math.isnan(value):
                f.write(f"{key} = None\n")

            # format strings properly
            elif isinstance(value, str):
                f.write(f'{key} = "{value}"\n')

            # everything else (numbers, booleans)
            else:
                f.write(f"{key} = {value}\n")

def clean_value(v):
    if isinstance(v, str):
        v = v.strip()

        if v.lower() == "true":
            return True
        if v.lower() == "false":
            return False
        if v.lower() == "none":
            return None

        try:
            if "." in v:
                return float(v)
            return int(v)
        except:
            return v

    return v

get_input_file_path()
config_df = pd.read_excel(_selected_file_path, sheet_name="config")
config_dict = {
    row["variable"]: clean_value(row["value"])
    for _, row in config_df.iterrows()
}
max_timesteps = config_dict.get("max_timesteps")
config_dict["operation_and_maintenance"] = 10000 * (
    max_timesteps / (24 * 4 * 365) if not np.isnan(max_timesteps) else 1 ## "nan" instead of None
)
write_config_py(config_dict)

# ==========================================================
# Select sheets from excel file
# ==========================================================
def select_sheets(label_text):
    """
    Opens a file dialog with the excel sheets in the selected excel file
    """

    def selection_list():
        selection.clear()
        for var in vars:
            if var.get() == 1:
                selection.append(sheet_list[vars.index(var)])
        root.destroy()

        

    file_path = get_input_file_path()
    df = pd.read_excel(file_path, sheet_name=None)
    sheet_list = [
        sheet for sheet in df.keys()
        if sheet != "config" and not str(sheet).startswith("_xlnm.")
    ]
    selection = []
    root = tk.Tk()
    root.title("Select Data Sheet(s)")
    root.geometry('500x250')

    # Add a label at the top
    label = tk.Label(root, text=label_text, font=("Arial", 12, "bold"))
    label.pack(side='top', pady=5)

    vars = [tk.IntVar() for sheet in sheet_list]
    btns = [tk.Checkbutton(root, text = sheet, variable = vars[sheet_list.index(sheet)],
                        onvalue=1, offvalue=0) for sheet in sheet_list]

    for btn in btns:
        btn.pack(side = 'top')

    sbtn = tk.Button(root, text = 'Apply & Exit', 
                        command = selection_list) 
    sbtn.pack(side = 'top')      

    root.mainloop()

    return(selection)


def _parse_excel_time(value):
    """
    Convert Excel/Pandas/string time values to a Python time object.
    """
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.time()

    if isinstance(value, datetime):
        return value.time()

    if isinstance(value, time):
        return value

    if isinstance(value, (int, float)):
        # Excel stores times as fractions of a day.
        try:
            base = pd.Timestamp("1899-12-30")
            return (base + pd.to_timedelta(float(value), unit="D")).time()
        except Exception:
            return None

    try:
        parsed = pd.to_datetime(str(value), format="%H:%M:%S", errors="coerce")
        return None if pd.isna(parsed) else parsed.time()
    except Exception:
        return None


def _normalize_cell_ref(ref: str) -> str:
    return ref.replace("$", "").upper()


def _evaluate_excel_formula_cell(worksheet, cell_ref: str, cache: dict):
    """
    Evaluate a small subset of Excel formulas needed by the input workbook.
    Supports direct values plus formulas using SUM/MIN/MAX and arithmetic.
    """
    cell_ref = _normalize_cell_ref(cell_ref)
    if cell_ref in cache:
        return cache[cell_ref]

    cell = worksheet[cell_ref]
    value = cell.value

    if value is None:
        cache[cell_ref] = None
        return None

    if isinstance(value, (int, float)):
        cache[cell_ref] = float(value)
        return cache[cell_ref]

    if isinstance(value, str) and value.startswith("="):
        formula = value[1:].strip()

        def cell_value(ref: str):
            return _evaluate_excel_formula_cell(worksheet, ref, cache) or 0.0

        def range_values(start_ref: str, end_ref: str):
            cells = worksheet[f"{_normalize_cell_ref(start_ref)}:{_normalize_cell_ref(end_ref)}"]
            values = []
            for row in cells:
                for item in row:
                    values.append(_evaluate_excel_formula_cell(worksheet, item.coordinate, cache) or 0.0)
            return values

        python_expr = re.sub(
            r"([A-Z]+\d+):([A-Z]+\d+)",
            lambda match: f'RANGE_VALUES("{match.group(1)}", "{match.group(2)}")',
            formula,
        )
        python_expr = re.sub(r"\bSUM\s*\(", "sum(", python_expr, flags=re.IGNORECASE)
        python_expr = re.sub(r"\bMIN\s*\(", "min(", python_expr, flags=re.IGNORECASE)
        python_expr = re.sub(r"\bMAX\s*\(", "max(", python_expr, flags=re.IGNORECASE)
        python_expr = re.sub(
            r"(?<![A-Z0-9_\"'])\b([A-Z]+\d+)\b(?![A-Z0-9_\"'])",
            lambda match: f'CELL("{match.group(1)}")',
            python_expr,
        )
        python_expr = python_expr.replace("^", "**")

        try:
            evaluated = eval(
                python_expr,
                {"__builtins__": {}},
                {"CELL": cell_value, "RANGE_VALUES": range_values, "sum": sum, "min": min, "max": max},
            )
        except Exception as exc:
            raise ValueError(f"Could not evaluate Excel formula '{value}' in cell {cell_ref}") from exc

        cache[cell_ref] = None if evaluated is None else float(evaluated)
        return cache[cell_ref]

    numeric = pd.to_numeric(value, errors="coerce")
    cache[cell_ref] = None if pd.isna(numeric) else float(numeric)
    return cache[cell_ref]


def _load_lkw_template_from_excel(file_path, sheet_name):
    """
    Read the LKW template with openpyxl so derived Excel columns can be
    recomputed even when the workbook has no cached formula results.
    """
    workbook = load_workbook(file_path, data_only=False, read_only=False)
    try:
        worksheet = workbook[sheet_name]
        formula_cache = {}
        rows = []

        for row_idx in range(3, worksheet.max_row + 1):
            time_value = _parse_excel_time(worksheet[f"A{row_idx}"].value)
            lkw_kw = _evaluate_excel_formula_cell(worksheet, f"E{row_idx}", formula_cache)
            if lkw_kw is None:
                total_kwh = _evaluate_excel_formula_cell(worksheet, f"D{row_idx}", formula_cache)
                if total_kwh is not None:
                    lkw_kw = 4 * total_kwh

            rows.append({"row_idx": row_idx, "time": time_value, "lkw_kW": lkw_kw})

        df = pd.DataFrame(rows)

        # Some templates leave the first quarter-hour unlabeled although the
        # derived power formula is present. Recover that midnight slot.
        first_data_row_missing_time = (
            not df.empty
            and df.iloc[0]["row_idx"] == 3
            and pd.isna(df.iloc[0]["time"])
            and pd.notna(df.iloc[0]["lkw_kW"])
        )
        if first_data_row_missing_time:
            df.loc[df["row_idx"] == 3, "time"] = time(0, 0)

        df = df.dropna(subset=["time", "lkw_kW"]).sort_values("time")
        return df.drop(columns=["row_idx"])
    finally:
        workbook.close()



# ==========================================================
# 10 → 15 minute conversion
# ==========================================================
def convert_to_15min(df, column_name="power_kW"):
    """
    Convert a 10-minute average power time series (kW)
    into a 15-minute average power time series (kW).

        1) Convert power → energy (kWh)
        2) Move energy to a 5-minute grid
        3) Aggregate energy to 15-minute blocks
        4) Convert energy back → power

    This guarantees energy conservation.
    """

    df = df.copy()
    df = df.set_index("timestamp")

    df["energy_kWh"] = df[column_name] * (10 / 60)

    # Upsample to 5-minute grid
    df_5 = df["energy_kWh"].resample("5min").asfreq()
    df_5 = df_5.ffill() / 2

    # Aggregate to 15-minute
    energy_15 = df_5.resample("15min").sum()

    power_15 = energy_15 / (15 / 60)

    result = power_15.reset_index()
    result.columns = ["timestamp", column_name]

    return result


def fill_missing_data(df):
    """
    expects: dataframe with timestamp of 15min intervalls

    Fill missing values:
    1. interpolate gaps up to 1h
    2. fill remaining gaps using previous day
    3. fill remaining gaps using next day
    does this for up to two weeks max
    """

    df = df.interpolate(limit=4)

    # fill from previous days
    for i in range(1, 15):  # up to two weeks
        df = df.fillna(df.shift(96 * i))

    # fill from following days
    for i in range(1, 15):
        df = df.fillna(df.shift(-96 * i))

    return df

# ==========================================================
# Generic trafo loader
# ==========================================================
def load_trafo(sheet_name):
    """
    Load and clean a transformer sheet.
    Returns:
        DataFrame with columns:
            timestamp (datetime)
            power_kW (float)
    """

    file_path = get_input_file_path()
    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=1
    )

    # Detect avg column automatically
    avg_columns = [col for col in df.columns if "-avg[W]" in col]
    if not avg_columns:
        raise ValueError(f"No avg column found in sheet {sheet_name}")
    avg_column = avg_columns[0]

    df = df[["Zeit", avg_column]]
    df.columns = ["timestamp", "power_W"]

    # Convert types
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df["power_W"] = pd.to_numeric(df["power_W"], errors="coerce")

    df = df.dropna()

    # Convert W → kW
    df["power_kW"] = df["power_W"] / 1000

    df = df[["timestamp", "power_kW"]]

    df = df.sort_values("timestamp")
    df = df.drop_duplicates(subset="timestamp", keep="first")

    df = df.set_index("timestamp")

    full_index = pd.date_range(
        start=df.index.min(),
        end=df.index.max(),
        freq="10min"
    )

    df = df.reindex(full_index)

    df["power_kW"] = df["power_kW"].interpolate(method="time")

    df = df.reset_index()
    df.columns = ["timestamp", "power_kW"]
    df_15 = convert_to_15min(df,"power_kW")
    return df_15


# ==========================================================
# Grid exchange (Trafo1 + Trafo2)
# ==========================================================
def load_grid_exchange(trafo_sheets):
    """
    Returns structured dataframe with:
        timestamp
        trafo1_kW
        trafo2_kW
        grid_exchange_kW
    """

    df = None
    trafo_cols = []

    for i, sheet in enumerate(trafo_sheets, start=1):
        trafo = load_trafo(sheet)

        col_name = f"trafo{i}_kW"
        trafo = trafo.rename(columns={"power_kW": col_name})

        trafo_cols.append(col_name)

        if df is None:
            df = trafo
        else:
            df = df.merge(trafo, on="timestamp", how="outer")

    # total grid exchange
    df["grid_exchange_kW"] = df[trafo_cols].sum(axis=1)

    return df[["timestamp", *trafo_cols, "grid_exchange_kW"]]


# ==========================================================
# EV demand LKW
# ==========================================================
def generate_lkw_profile(file_path=None, year=None, sheet_name=None):
    """
    Generate full-year charging profile for 2025 in 15-minute intervall from single example day data.

    Assumptions:
        - Sheet contains one typical 15-min weekday
        - Charging only Monday–Friday
        - Saturday and Sunday = 0
        - no seasonality
    """

    if file_path is None:
        file_path = get_input_file_path()

    if year is None:
        print("no year detected for charging profile generation")
        return
    
    if sheet_name is None:
        print("no worksheet selected for charging profile generation")
        return

    df = _load_lkw_template_from_excel(file_path, sheet_name)

    if len(df) != 96:
        raise ValueError(
            f"LKW profile in sheet '{sheet_name}' must contain 96 valid 15-min rows after parsing. "
            f"Found {len(df)} rows."
        )

    daily_profile = df["lkw_kW"].values

    # Create full-year 15-min index
    start = pd.Timestamp(f"{year}-01-01 00:00:00")
    end = pd.Timestamp(f"{year}-12-31 23:45:00")
    full_index = pd.date_range(start=start, end=end, freq="15min")

    result = pd.DataFrame({"timestamp": full_index})
    result["weekday"] = result["timestamp"].dt.weekday  # Monday=0, Sunday=6

    # Initialize with 0
    result["lkw_kW"] = 0.0

    # Apply profile only on weekdays
    weekday_mask = result["weekday"] < 5

    # Repeat daily profile for number of weekdays
    weekday_indices = result[weekday_mask].index
    num_weekdays = len(weekday_indices) // 96

    repeated_profile = np.tile(daily_profile, num_weekdays)

    result.loc[weekday_mask, "lkw_kW"] = repeated_profile

    result = result.drop(columns="weekday")

    return result

# ==========================================================
# EV demand Zustellung
# ==========================================================
def generate_zustellung_profile(file_path=None, year=None, sheet_name=None):
    """
    Generate full-year 15-min Zustellung load profile.

    - Winter profile from Excel
    - Mixed resolution (hourly + 15-min) harmonized to 15-min
    - Weekday-dependent (Mon–Sat)
    - Sunday = 0
    - Summer months reduced by 40%
    """

    if file_path is None:
        file_path = get_input_file_path()

    if year is None:
        print("no year detected for delivery profile generation")
        return
    
    if sheet_name is None:
        print("no worksheet selected for delivery profile generation")
        return

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=0
    )

    # Remove empty Excel columns
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    # Rename first column to "time"
    df = df.rename(columns={df.columns[0]: "time"})

    # Robust time parsing
    def parse_time(value):
        try:
            if isinstance(value, pd.Timestamp):
                return value.time()
            return pd.to_datetime(value, format="%H:%M:%S").time()
        except:
            return None

    df["time"] = df["time"].apply(parse_time)
    df = df.dropna(subset=["time"])

    # Convert weekday columns to numeric
    weekday_cols = df.columns.drop("time")
    for col in weekday_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.sort_values("time").reset_index(drop=True)

    # harmonize to 15 minute resolution
    df["timestamp"] = (
        pd.to_datetime("1900-01-01")
        + pd.to_timedelta(df["time"].astype(str))
    )

    df = df.set_index("timestamp")
    df = df.drop(columns=["time"])

    full_index = pd.date_range(
        start="1900-01-01 00:00:00",
        end="1900-01-01 23:45:00",
        freq="15min"
    )

    df_15 = df.reindex(full_index).ffill()

    df_15 = df_15.reset_index().rename(columns={"index": "timestamp"})
    df_15["time"] = df_15["timestamp"].dt.time
    df_15 = df_15.drop(columns=["timestamp"])

    # full year timestamp
    start = pd.Timestamp(f"{year}-01-01 00:00:00")
    end = pd.Timestamp(f"{year}-12-31 23:45:00")

    full_year_index = pd.date_range(start=start, end=end, freq="15min")

    result = pd.DataFrame({"timestamp": full_year_index})
    result["weekday_num"] = result["timestamp"].dt.weekday
    result["time"] = result["timestamp"].dt.time

    # mapping weekday profiles
    lookup = {}
    for day in df_15.columns:
        if day != "time":
            lookup[day] = dict(zip(df_15["time"], df_15[day]))

    weekday_map = {
        0: "Montag",
        1: "Dienstag",
        2: "Mittwoch",
        3: "Donnerstag",
        4: "Freitag",
        5: "Samstag"
    }

    result["zustellung_kW"] = 0.0

    for wd_num, wd_name in weekday_map.items():
        mask = result["weekday_num"] == wd_num
        result.loc[mask, "zustellung_kW"] = (
            result.loc[mask, "time"].map(lookup[wd_name])
        )

    # Sunday remains 0 automatically

    # apply summer reduction of -40%
    summer_months = [4, 5, 6, 7, 8, 9]  # April–September

    summer_mask = result["timestamp"].dt.month.isin(summer_months)
    result.loc[summer_mask, "zustellung_kW"] *= 0.6

    result = result[["timestamp", "zustellung_kW"]]

    return result



# ==========================================================
# config helpers
# ==========================================================
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

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import tkinter as tk
from tkinter import filedialog

## Will be used to process the input data before feeding the data into the optimization framework

file_path = filedialog.askopenfilename(initialdir = "01-INPUT-DATA",
                                          title = "Select Input Data",
                                          filetypes = [("Excel files", ".xlsx .xls")])
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

        

    df = pd.read_excel(file_path,sheet_name = None)
    sheet_list = list(df.keys())
    selection = []
    root = tk.Tk()
    root.title("Select Trafo Data Sheet(s)")
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
def generate_lkw_profile(file_path=file_path, year=2025):
    """
    Generate full-year charging profile for 2025 in 15-minute intervall from single example day data.

    Assumptions:
        - Sheet contains one typical 15-min weekday
        - Charging only Monday–Friday
        - Saturday and Sunday = 0
        - no seasonality
    """

    df = pd.read_excel(
        file_path,
        sheet_name="Zubau_LKW_Aug2025",
        header=1
    )

    # Find exact Total kW column
    power_col = None
    for col in df.columns:
        if str(col).strip() == "Total kW":
            power_col = col
            break

    if power_col is None:
        raise ValueError("Exact column 'Total kW' not found.")

    time_col = df.columns[0]
    df = df[[time_col, power_col]]
    df.columns = ["time", "lkw_kW"]

    df["time"] = pd.to_datetime(df["time"], format="%H:%M:%S", errors="coerce")
    df["lkw_kW"] = pd.to_numeric(df["lkw_kW"], errors="coerce")

    df = df.dropna().sort_values("time")

    if len(df) != 96:
        raise ValueError("LKW profile must contain 96 rows (15-min full weekday).")

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
def generate_zustellung_profile(file_path=file_path, year=2026):
    """
    Generate full-year 15-min Zustellung load profile.

    - Winter profile from Excel
    - Mixed resolution (hourly + 15-min) harmonized to 15-min
    - Weekday-dependent (Mon–Sat)
    - Sunday = 0
    - Summer months reduced by 40%
    """

    df = pd.read_excel(
        file_path,
        sheet_name="Zubau_Zustellung_Oct2026",
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


import pandas as pd
import numpy as np
import csv
import tkinter as tk
from tkinter import filedialog

## Will be used to process the input data before feeding the data into the optimization framework

#file_path = "01-INPUT-DATA/Vétroz/20251023_Standortdaten_Vétroz.xlsx"
file_path = filedialog.askopenfilename(initialdir = "01-INPUT-DATA",
                                          title = "Select Input Data",
                                          filetypes = [("Excel files", ".xlsx .xls")])
# ==========================================================
# Select sheets from excel file
# ==========================================================
def select_sheets():

    def selection_list():
        for var in vars:
            if var.get() == 1:
                selection.append(sheet_list[vars.index(var)])

    df = pd.read_excel(file_path,sheet_name = None)
    sheet_list = list(df.keys())
    selection = []
    root = tk.Tk()
    root.title("Select Trafo Data Sheet(s)")
    root.geometry = ('100x100')

    vars = [tk.IntVar() for sheet in sheet_list]
    btns = [tk.Checkbutton(root, text = sheet, variable = vars[sheet_list.index(sheet)],
                        onvalue=1, offvalue=0) for sheet in sheet_list]

    for btn in btns:
        btn.pack(side = 'top')

    sbtn = tk.Button(root, text = 'Apply', 
                        command = selection_list) 
    sbtn.pack(side = 'top')      

    xbtn = tk.Button(root, text = 'Exit', 
                        command = root.destroy) 
    xbtn.pack(side = 'top')      
    root.mainloop()

    return(selection)

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

    return df

# ==========================================================
# Grid exchange (Trafo1 + Trafo2)
# ==========================================================
def load_grid_exchange():
    """
    Returns structured dataframe with:
        timestamp
        trafo1_kW
        trafo2_kW
        grid_exchange_kW
    """

    trafo1 = load_trafo("2024_Verbrauch_Trafo1")
    trafo2 = load_trafo("2024_Verbrauch_Trafo2")

    df = trafo1.merge(
        trafo2,
        on="timestamp",
        suffixes=("_t1", "_t2")
    )

    df = df.rename(columns={
        "power_kW_t1": "trafo1_kW",
        "power_kW_t2": "trafo2_kW"
    })

    df["grid_exchange_kW"] = df["trafo1_kW"] + df["trafo2_kW"]

    return df[["timestamp", "trafo1_kW", "trafo2_kW", "grid_exchange_kW"]]

# ==========================================================
# 10 → 15 minute conversion
# ==========================================================
def convert_to_15min(df, column_name):
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
# Check
# ==========================================================

##TEST##

selection = select_sheets()

print(selection)

##TEST##
'''
if __name__ == "__main__":

    print("=== TRAFO CHECK ===")
    trafo1 = load_trafo("2024_Verbrauch_Trafo1")
    trafo2 = load_trafo("2024_Verbrauch_Trafo2")

    print("Trafo 1 rows:", len(trafo1))
    print("Trafo 2 rows:", len(trafo2))

    print("\nTime diff Trafo1:")
    print(trafo1["timestamp"].diff().value_counts())

    print("\n=== GRID EXCHANGE CHECK ===")
    grid = load_grid_exchange()

    print("Grid rows:", len(grid))
    print(grid.head())

    print("\nEnergy consistency check:")
    energy_10 = (grid["grid_exchange_kW"] * (10/60)).sum()

    grid_15 = convert_to_15min(grid, "grid_exchange_kW")
    energy_15 = (grid_15["grid_exchange_kW"] * (15/60)).sum()

    print("Total energy 10-min:", energy_10)
    print("Total energy 15-min:", energy_15)

    print("\n15-min spacing check:")
    print(grid_15["timestamp"].diff().value_counts())
'''

# ==========================================================
# Visualizing the data
# ==========================================================
""" import matplotlib.pyplot as plt
import matplotlib.dates as mdates

if __name__ == "__main__":

    grid = load_grid_exchange()

    start_date = "2024-02-10"
    end_date   = "2024-02-11"

    data = grid[
        (grid["timestamp"] >= start_date) &
        (grid["timestamp"] < end_date)
    ].copy()

    plt.figure()

    plt.plot(data["timestamp"], data["trafo1_kW"])
    plt.plot(data["timestamp"], data["trafo2_kW"])
    plt.plot(data["timestamp"], data["grid_exchange_kW"])

    # Dynamic axis formatting
    if (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days <= 1:
        # Single day
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.gca().xaxis.set_major_locator(mdates.HourLocator(interval=1))
    else:
        # Multi-day
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d-%m'))
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))

    plt.xticks(rotation=45)

    plt.xlabel("Time")
    plt.ylabel("Power (kW)")
    plt.title(f"Trafo1, Trafo2 & Grid Exchange\n{start_date} to {end_date}")
    plt.legend(["Trafo1", "Trafo2", "Grid Exchange"])

    plt.tight_layout()
    plt.show() """
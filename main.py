import pandas as pd
import numpy as np

import optimization as opt
import data_preprocessing as dpp
import results_processing as rp

from config_loader import load_config, print_config


def main():
    cfg = load_config("config.json")
    print_config(cfg)

    # Beispiel: Config-Werte verwenden
    solver = cfg["run"]["solver"]
    timestep = cfg["run"]["timestep_minutes"]
    excel_path = cfg["inputs"]["excel_path"]

    # Beispiel: Excel einlesen (je nach euren Modulen könnt ihr das auch in dpp packen)
    load_sheet = cfg["inputs"].get("load_sheet", "load")
    load_df = pd.read_excel(excel_path, sheet_name=load_sheet)

    # Optional: timestamp sauber parsen
    if "timestamp" in load_df.columns:
        load_df["timestamp"] = pd.to_datetime(load_df["timestamp"])

    print("Loaded load_df head:")
    print(load_df.head())

    # Dann eure Pipeline:
    # preprocessed = dpp.preprocess(load_df, cfg)  # falls ihr cfg durchreichen wollt
    # model = opt.build_model(preprocessed, cfg)
    # results = opt.solve(model, solver=solver)
    # rp.export_results(results, cfg)

if __name__ == "__main__":
    main()
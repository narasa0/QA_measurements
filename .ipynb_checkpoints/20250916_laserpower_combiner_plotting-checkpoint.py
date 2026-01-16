#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 16 13:49:32 2025

@author: narasa01
"""

import pandas as pd
import matplotlib.pyplot as plt
import glob
import os
import re

# --- Function to extract "Result table values" into DataFrame ---
def read_power_instruction_table(file_path):
    with open(file_path, "r") as f:
        lines = f.readlines()
    
    # Find start of "Result table values"
    start_idx = None
    for i, line in enumerate(lines):
        if "Result table values" in line:
            start_idx = i + 1
            break
    if start_idx is None:
        raise ValueError(f"No 'Result table values' found in {file_path}")
    
    # Extract header
    header = lines[start_idx].strip().split(";")
    
    # Collect rows until "time" or empty line
    data = []
    for line in lines[start_idx+1:]:
        if line.strip().lower().startswith("time"):
            break
        if not line.strip():
            break
        data.append(line.strip().split(";"))
    
    df = pd.DataFrame(data, columns=header)

    # Rename column
    df = df.rename(columns={"power_instruction": "power_percentage_values"})

    # Convert columns to numeric where possible
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except (ValueError, TypeError):
            pass
    return df


# --- Function to combine multiple files of same wavelength ---
def combine_group(dfs_dict, wavelength):
    combined = None
    for fname, df in dfs_dict.items():
        date = fname.split("_")[0]  # e.g. "07-25"
        
        # Rename wavelength-specific columns
        cols = [c for c in df.columns if c != "power_percentage_values"]
        rename_dict = {c: f"{date}_{c}" for c in cols}
        df_renamed = df.rename(columns=rename_dict)
        
        if combined is None:
            combined = df_renamed
        else:
            combined = pd.merge(combined, df_renamed, 
                                on="power_percentage_values", how="outer")
    
    if combined is not None:
        combined = (
            combined.drop_duplicates(subset=["power_percentage_values"])
                    .sort_values("power_percentage_values")
        )
    return combined


# --- Function to plot wavelength data ---
def plot_wavelength_data(combined_all, save_folder="plots"):
    os.makedirs(save_folder, exist_ok=True)

    for wl, df in combined_all.items():
        plt.figure(figsize=(8, 6))
        
        # Extract months from column names
        month_map = {}
        for col in df.columns:
            if col == "power_percentage_values":
                continue
            if "_power" in col:
                month = col.split("_")[0]   # e.g. "07-25"
                month_num = int(month.split("-")[0])
                month_map[month] = month_num
        
        # Sort months numerically
        for month in sorted(month_map, key=lambda x: month_map[x]):
            power_col = [c for c in df.columns if c.startswith(month) and "_power" in c][0]
            error_col = [c for c in df.columns if c.startswith(month) and "_error" in c][0]
            
            plt.errorbar(
                df["power_percentage_values"], 
                df[power_col], 
                yerr=df[error_col].abs(), 
                marker="o", capsize=3, 
                label=f"{month}"
            )
        
        plt.title(f"Laser Power Calibration - {wl}nm")
        plt.xlabel("Power Percentage Values")
        plt.ylabel("Measured Power (mW)")
        plt.legend(title="Month")
        plt.grid(True, linestyle="--", alpha=0.5)
        plt.tight_layout()
        
        # Save plot as PNG
        plot_file = os.path.join(save_folder, f"laser_power_{wl}nm.png")
        plt.savefig(plot_file, dpi=300)
        plt.close()
        
        print(f"✅ Saved plot for {wl}nm → {plot_file}")


# --- Main workflow ---
def main(input_folder, output_excel, plot_folder):
    # Collect all CSVs
    files = glob.glob(os.path.join(input_folder, "*.csv"))
    if not files:
        print("⚠️ No CSV files found in folder.")
        return
    
    # Load all DataFrames
    dfs = {os.path.basename(file): read_power_instruction_table(file) for file in files}

    # Group automatically by wavelength
    grouped = {}
    for fname, df in dfs.items():
        match = re.search(r"_(\d+)\.csv", fname)
        if match:
            wl = match.group(1)
            grouped.setdefault(wl, {})[fname] = df

    # Combine
    combined_all = {wl: combine_group(dfs_dict, wl) for wl, dfs_dict in grouped.items()}

    # Export results
    try:
        import openpyxl
        if combined_all:  # Only proceed if we actually have data
            with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
                for wl in sorted(combined_all.keys(), key=lambda x: int(x)):
                    df = combined_all[wl]
                    if df is not None and not df.empty:
                        df.to_excel(writer, sheet_name=f"{wl}nm", index=False)
            print(f"✅ Exported all wavelengths (sorted) to {output_excel}")
        else:
            print("⚠️ No data found to export.")
    except ImportError:
        print("⚠️ openpyxl not installed. Exporting as separate CSV files instead...")
        for wl in sorted(combined_all.keys(), key=lambda x: int(x)):
            df = combined_all[wl]
            if df is not None and not df.empty:
                csv_file = f"combined_{wl}nm.csv"
                df.to_csv(csv_file, index=False)
                print(f"✅ Saved {csv_file}")

    # Plot data
    plot_wavelength_data(combined_all, save_folder=plot_folder)


# --- Run script ---
if __name__ == "__main__":
    main(
        input_folder="/Users/narasa01/Documents/data/QA/Laser_power/microscopes/Leica_STED",  # <-- change this to your CSV folder
        output_excel="/Users/narasa01/Documents/data/QA/Laser_power/microscopes/Leica_STED/outputs/combined_power_data.xlsx",
        plot_folder="/Users/narasa01/Documents/data/QA/Laser_power/microscopes/Leica_STED/outputs"
    )

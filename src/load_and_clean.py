import pandas as pd
import numpy as np

# --- Define File Paths ---

# Sector-level files
file_nbr_processed = "Merged_NBR_Uti_Study_LTE_Data-Ardebil_processed.csv"
file_sector_kpis = "Ardebil sector level.csv"
file_lte_physical = "LTE.csv" # Physical sector data

# Cell-level configuration files
file_cell_config_main = "CEll Configuration-ardebil.csv"
file_interfreq_config = "INTERFREQ ConfigurationReport_2025_04_29_09_30_42_998.csv"
file_cell_static_params = "LST CELL+Display static parameters of cells-ardebil-.csv"
file_mo_report_intra_freq_ncell = "MoReport_EutranIntraFreqNCell-ardebil.csv" # Previously df_intra_freq_ncell
file_pdsch_config = "PDSCH Pwer config Cell_2025_05_25-ardebil.csv"
file_eutran_interfreq_ncell = "EUTRANINTERFREQNCELL.csv" # New cell config file

# Cell-level hourly KPI file (single merged file)
file_hourly_cell_kpis_merged = "Ardebil cell levelKPI Hourly.csv"

# Dictionary to store loaded dataframes
dataframes = {}

def load_and_inspect(df_name, file_path, is_time_series=False, time_column='time'):
    """Helper function to load a CSV, inspect it, and convert time column."""
    print(f"\n--- Loading and Inspecting: {file_path} (as {df_name}) ---")
    try:
        try:
            df = pd.read_csv(file_path)
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='latin1')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='iso-8859-1')
        
        print("Head:")
        print(df.head(3))
        
        if is_time_series and time_column in df.columns:
            df[time_column] = pd.to_datetime(df[time_column], errors='coerce')
            if df[time_column].isnull().any():
                print(f"Warning: Some '{time_column}' values in {file_path} could not be parsed and are NaT.")
            print(f"'{time_column}' column converted to datetime.")
            
        print("\nInfo:")
        df.info()
        dataframes[df_name] = df
    except FileNotFoundError:
        print(f"Error: '{file_path}' not found.")
        dataframes[df_name] = pd.DataFrame()
    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        dataframes[df_name] = pd.DataFrame()
    print("\n" + "="*50)

# --- Load Sector-Level Files ---
load_and_inspect("df_nbr_processed", file_nbr_processed)
load_and_inspect("df_sector_level_kpis", file_sector_kpis, is_time_series=True)
load_and_inspect("df_lte_physical", file_lte_physical)

# --- Load Cell-Level Configuration Files ---
load_and_inspect("df_cell_config_main", file_cell_config_main)
load_and_inspect("df_interfreq_config", file_interfreq_config)
load_and_inspect("df_cell_static_params", file_cell_static_params)
load_and_inspect("df_mo_report_intra_freq_ncell", file_mo_report_intra_freq_ncell) # Renamed df variable for clarity
load_and_inspect("df_pdsch_config", file_pdsch_config)
load_and_inspect("df_eutran_interfreq_ncell", file_eutran_interfreq_ncell) # New file

# --- Load Merged Cell-Level Hourly KPI Data ---
load_and_inspect("df_hourly_cell_kpis", file_hourly_cell_kpis_merged, is_time_series=True)


print("\nSummary of loaded DataFrames (access them using dataframes['df_name']):")
for name, df_content in dataframes.items():
    if not df_content.empty:
        print(f"- {name}: {df_content.shape[0]} rows, {df_content.shape[1]} columns")
    else:
        print(f"- {name}: Empty (file not found or error during loading)")
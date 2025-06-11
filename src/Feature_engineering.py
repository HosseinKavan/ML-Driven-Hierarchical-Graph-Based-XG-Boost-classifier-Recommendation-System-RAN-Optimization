import pandas as pd
import numpy as np
import re

# --- Helper Function for Standardization (from previous discussions) ---
def standardize_sector_id(sector_id_series):
    """Standardizes sector IDs to a common format (e.g., LH..., LT...)."""
    if not isinstance(sector_id_series, pd.Series):
        sector_id_series = pd.Series(sector_id_series)
    temp_series = sector_id_series.astype(str)
    standardized_ids = []
    for sid in temp_series:
        if pd.isna(sid) or sid.lower() == 'nan':
            standardized_ids.append(np.nan)
            continue
        sid_str = sid.strip().upper()
        if re.match(r"^[HLT][A-Z0-9]{3,}[A-Z]$", sid_str):
            if sid_str.startswith('H') and not sid_str.startswith('LH'):
                standardized_ids.append('L' + sid_str)
            elif sid_str.startswith('T') and not sid_str.startswith('LT'):
                standardized_ids.append('L' + sid_str)
            else:
                standardized_ids.append(sid_str)
        else:
            standardized_ids.append(sid_str)
    return pd.Series(standardized_ids, index=sector_id_series.index)

# --- 1. Load the required files ---
try:
    # The main sector-level file to which we will add cell names
    df_sector = pd.read_csv("Merged_NBR_Uti_Study_LTE_Data-Ardebil_processed.csv")
    
    # The cell configuration file that contains the mapping from sector to cells
    df_cell_config = pd.read_csv("CEll Configuration-ardebil.csv")
    print("Files loaded successfully.")

except FileNotFoundError as e:
    print(f"Error loading files: {e}. Please ensure both CSV files are in the correct directory.")
    exit()

# --- 2. Standardize Sector IDs in both DataFrames for accurate matching ---
print("Standardizing sector identifiers...")
# Standardize the 'Source' column in the sector data
df_sector['Canonical_Sector_ID'] = standardize_sector_id(df_sector['Source'])

# Standardize the 'Sector' column in the cell config data
df_cell_config['Canonical_Sector_ID'] = standardize_sector_id(df_cell_config['Sector'])

# --- 3. Group Cell Names by Sector ---
# From the cell configuration file, create a list of all cell names for each sector.
print("Grouping cell names by sector...")
cell_map = df_cell_config.groupby('Canonical_Sector_ID')['Cell Name'].apply(list).reset_index()
cell_map.rename(columns={'Cell Name': 'Cell_Name_List'}, inplace=True)

print("\nSample of Sector-to-Cell mapping:")
print(cell_map.head())

# --- 4. Merge the Cell Name Lists into the Sector DataFrame ---
print("\nMerging cell name lists into the main sector data file...")
# Use a left merge to ensure all original sectors from df_sector are kept
df_sector_with_cells = pd.merge(df_sector, cell_map, on='Canonical_Sector_ID', how='left')

# --- 5. Expand the Cell Name Lists into Separate Columns ---
# Check if the 'Cell_Name_List' column exists and has content
if 'Cell_Name_List' in df_sector_with_cells.columns and df_sector_with_cells['Cell_Name_List'].notna().any():
    print("Expanding cell name lists into separate columns (Cell Name 1, Cell Name 2, etc.)...")
    # This creates a new DataFrame with columns 0, 1, 2, ... for each cell in the list
    cell_name_cols = df_sector_with_cells['Cell_Name_List'].apply(pd.Series)
    
    # Rename the columns of the new DataFrame
    cell_name_cols.columns = [f"Cell Name {i+1}" for i in cell_name_cols.columns]
    
    # Join the new cell name columns back to our main DataFrame
    df_final = pd.concat([df_sector_with_cells, cell_name_cols], axis=1)
    
    # Drop the intermediate list column and the temporary canonical ID
    df_final.drop(columns=['Cell_Name_List', 'Canonical_Sector_ID'], inplace=True)
    
    print("\n--- Final DataFrame Head with Cell Names Added ---")
    # Display relevant original columns and the new cell name columns
    display_cols = ['Source'] + [col for col in df_final.columns if col.startswith('Cell Name')]
    print(df_final[display_cols].head())
    
    # Save the final result to a new CSV file
    output_filename = "Sectors_with_Cell_Names.csv"
    df_final.to_csv(output_filename, index=False)
    print(f"\nSuccessfully saved the result to '{output_filename}'")

else:
    print("\nWarning: No matching cells were found, or the 'Cell_Name_List' column was empty. No cell name columns were added.")
    print("Final DataFrame Head (unchanged):")
    print(df_sector_with_cells.head())
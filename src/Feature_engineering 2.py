import pandas as pd
import numpy as np
import re

# --- Helper Function for Standardization ---
def standardize_sector_id(sector_id_series):
    """Standardizes sector IDs to a common format (e.g., LH..., LT...)."""
    if not isinstance(sector_id_series, pd.Series):
        sector_id_series = pd.Series(sector_id_series)
    # Ensure series is treated as string for manipulation
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
    # The main sector-level file containing the Target NBRx columns
    df_sector = pd.read_csv("Merged_NBR_Uti_Study_LTE_Data-Ardebil_processed.csv")
    
    # The cell configuration file for the lookup
    df_cell_config = pd.read_csv("CEll Configuration-ardebil.csv")
    print("Files loaded successfully.")

except FileNotFoundError as e:
    print(f"Error loading files: {e}. Please ensure both CSV files are in the correct directory.")
    exit()

# --- 2. Create the Sector-to-Cells Lookup Map ---
print("Creating a lookup map of sectors to their corresponding cell names...")
# Standardize the 'Sector' column in the cell config data for accurate mapping
df_cell_config['Canonical_Sector_ID'] = standardize_sector_id(df_cell_config['Sector'])

# Group by the standardized sector ID and aggregate cell names into a list
cell_map_df = df_cell_config.groupby('Canonical_Sector_ID')['Cell Name'].apply(list)
# Convert the resulting Series to a dictionary for fast lookups
sector_to_cells_dict = cell_map_df.to_dict()

print("Lookup map created successfully.")
print("Example entry from map -> 'LH1000XB':", sector_to_cells_dict.get('LH1000XB', 'Not Found'))

# --- 3. Find and Map Cells for each 'Targets NBRx' column ---
# Identify all columns that start with 'Targets NBR'
target_nbr_cols = [col for col in df_sector.columns if col.startswith('Targets NBR')]

print(f"\nFound {len(target_nbr_cols)} Target NBR columns to process: {target_nbr_cols}")

for nbr_col in target_nbr_cols:
    if nbr_col in df_sector.columns:
        print(f"Processing column: {nbr_col}")
        
        # Standardize the sector names in the current Target NBR column
        standardized_nbr_sectors = standardize_sector_id(df_sector[nbr_col])
        
        # Create the new column name for the cell lists
        new_cell_list_col = f"{nbr_col} Cells"
        
        # Use the .map() function with our dictionary to find the list of cells for each NBR sector
        df_sector[new_cell_list_col] = standardized_nbr_sectors.map(sector_to_cells_dict)

print("\nFinished mapping cells to all Target NBR columns.")

# --- 4. Display Results and Save ---
print("\n--- Final DataFrame Head with Target NBR Cells Added ---")
# Create a list of columns to display for verification
# We'll show the first two NBR columns and their corresponding new cell list columns
cols_to_display = []
for i in range(1, 3): # Display for NBR1 and NBR2
    if f'Targets NBR{i}' in df_sector.columns:
        cols_to_display.append(f'Targets NBR{i}')
    if f'Targets NBR{i} Cells' in df_sector.columns:
        cols_to_display.append(f'Targets NBR{i} Cells')

# Also show the source sector
if 'Source' in df_sector.columns:
    cols_to_display.insert(0, 'Source')

# Print head of the selected columns
if cols_to_display:
    print(df_sector[cols_to_display].head())
else:
    print("Could not find Target NBR columns to display.")

# Save the final result to a new CSV file
output_filename = "Sectors_with_Target_NBR_Cells.csv"
df_sector.to_csv(output_filename, index=False)
print(f"\nSuccessfully saved the result to '{output_filename}'")
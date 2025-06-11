import pandas as pd
import numpy as np

# Load the CSV file
file_path = "Merged_NBR_Uti_Study_LTE_Data-Ardebil.csv"
try:
    df = pd.read_csv(file_path)
    print("Original DataFrame head:")
    print(df.head())
    print("\nOriginal DataFrame info:")
    df.info()
except Exception as e:
    print(f"Error loading file {file_path}: {e}")
    # Exit if file loading fails
    exit()

# Columns identified as potentially containing comma-separated listscd dat
# You might need to adjust this list based on exact file content and requirements
columns_to_split = [
    'Targets',
    'Utilizations',
    'UDCLIs',
    'Percentages',
    'Distances',
    'High Risk', # Note: This might create columns with mixed data types if some rows have N/A or different structures
    'Low Risk',  # Similar to High Risk
    'Suitability Score',
    'NBRs Rank'
]

# Create a new DataFrame to hold the processed data
df_processed = df.copy()

for col_name in columns_to_split:
    if col_name in df_processed.columns:
        print(f"\nProcessing column: {col_name}")
        
        # Convert column to string type to ensure .str accessor works and handle potential NaNs gracefully
        # NaNs will become string 'nan' which usually doesn't split, or splits into one item ['nan']
        # If you want NaNs to result in NaN values in all split columns, specific handling might be needed
        # before or after the split. For now, let's ensure they don't break the split.
        # A common way is to fillna with an empty string before splitting if NaNs should result in empty split columns.
        df_processed[col_name] = df_processed[col_name].fillna('').astype(str)

        # Split the column by comma, creating new columns
        # The .str.strip() on each element after split can clean up leading/trailing spaces
        split_data = df_processed[col_name].str.split(',', expand=True)
        
        # Apply strip to each element in the new columns
        for split_col in split_data.columns:
            split_data[split_col] = split_data[split_col].str.strip()
            # Replace empty strings that might result from stripping whitespace-only entries or from original empty parts
            # with NaN or None if that's preferred over empty strings. For now, they remain as empty strings or original values.
            # If a value was genuinely NaN before fillna('').astype(str), it became 'nan'.
            # If it was empty string, it remains empty.
            # If it was ' , ', it might become [' ', ' '] then stripped to ['', ''].
            # Consider replacing empty strings with np.nan if that's more appropriate for subsequent analysis:
            # split_data[split_col] = split_data[split_col].replace(r'^\s*$', np.nan, regex=True)


        # Generate new column names
        new_col_names = [f"{col_name} NBR{i+1}" for i in range(split_data.shape[1])]
        split_data.columns = new_col_names
        
        # Concatenate new columns to the DataFrame
        df_processed = pd.concat([df_processed, split_data], axis=1)
        
        # Drop the original column
        df_processed.drop(col_name, axis=1, inplace=True)
        print(f"Created columns: {', '.join(new_col_names)}")
    else:
        print(f"Column '{col_name}' not found in the DataFrame.")

print("\nProcessed DataFrame head:")
print(df_processed.head())
print("\nProcessed DataFrame info:")
df_processed.info()

# Save the processed DataFrame to a new CSV file
output_file_path = "Merged_NBR_Uti_Study_LTE_Data-Ardebil_processed.csv"
try:
    df_processed.to_csv(output_file_path, index=False)
    print(f"\nProcessed data successfully saved to {output_file_path}")
except Exception as e:
    print(f"\nError saving processed data to {output_file_path}: {e}")
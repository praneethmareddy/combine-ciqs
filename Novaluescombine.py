import pandas as pd
import glob
import os

# Folder containing CIQ Excel files
ciq_folder = "./ciq_files/"

# Get all Excel files in the folder
ciq_files = glob.glob(os.path.join(ciq_folder, "*.xlsx"))

# Separate DataFrames for row-wise and column-wise structures
row_wise_data = []
column_wise_data = []

for file in ciq_files:
    df = pd.read_excel(file)
    
    # Check if CIQ contains column-wise cell structure
    if any("cellid1" in col.lower() for col in df.columns) and any("cell1type" in col.lower() for col in df.columns):
        column_wise_data.append(df)
    else:
        row_wise_data.append(df)

# Merge data for row-wise and column-wise structures
if row_wise_data:
    row_wise_df = pd.concat(row_wise_data, ignore_index=True)
    row_wise_path = os.path.join(ciq_folder, "master_row_wise_ciq.xlsx")
    row_wise_df.to_excel(row_wise_path, index=False)
    print(f"Row-wise Master CIQ saved to {row_wise_path}")

if column_wise_data:
    column_wise_df = pd.concat(column_wise_data, ignore_index=True)
    column_wise_path = os.path.join(ciq_folder, "master_column_wise_ciq.xlsx")
    column_wise_df.to_excel(column_wise_path, index=False)
    print(f"Column-wise Master CIQ saved to {column_wise_path}")

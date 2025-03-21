import pandas as pd
import glob
import os

# Folder containing CIQ Excel files
ciq_folder = "./ciq_files/"

# Get all Excel files in the folder
ciq_files = glob.glob(os.path.join(ciq_folder, "*.xlsx"))

# Separate column names for row-wise and column-wise structures
row_wise_columns = set()
column_wise_columns = set()

for file in ciq_files:
    df = pd.read_excel(file)
    
    # Check if CIQ contains column-wise cell structure
    if any("cellid1" in col.lower() for col in df.columns) and any("cell1type" in col.lower() for col in df.columns):
        column_wise_columns.update(df.columns)
    else:
        row_wise_columns.update(df.columns)

# Ensure master column structure for both row-wise and column-wise formats
row_wise_columns.update(["siteid", "neid", "cellid", "celltype", "xtype", "col1", "col2"])
column_wise_columns.update(["siteid", "neid", "cellid1", "celltype1", "cellid2", "celltype2", "xtype", "col1", "col2"])

# Create DataFrames with only column names
row_wise_df = pd.DataFrame(columns=sorted(row_wise_columns))
column_wise_df = pd.DataFrame(columns=sorted(column_wise_columns))

# Save the master templates for row-wise and column-wise formats
row_wise_path = os.path.join(ciq_folder, "master_row_wise_ciq.xlsx")
column_wise_path = os.path.join(ciq_folder, "master_column_wise_ciq.xlsx")

row_wise_df.to_excel(row_wise_path, index=False)
column_wise_df.to_excel(column_wise_path, index=False)

print(f"Row-wise Master CIQ saved to {row_wise_path}")
print(f"Column-wise Master CIQ saved to {column_wise_path}")

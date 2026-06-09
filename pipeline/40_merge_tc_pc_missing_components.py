import pandas as pd

from config import TC_PC_MISSING_REPORT, TC_PC_MERGED_FILE

input_file = TC_PC_MISSING_REPORT
output_file = TC_PC_MERGED_FILE

# Read sheets
tc_df = pd.read_excel(input_file, sheet_name="TC_Missing")
pc_df = pd.read_excel(input_file, sheet_name="PC_Missing")

# Rename columns for clarity
tc_df = tc_df.rename(columns={"Status": "TC_Status"})
pc_df = pc_df.rename(columns={"Status": "PC_Status"})

# Keep only required columns
tc_df = tc_df[["TRCID", "TC_Status"]]
pc_df = pc_df[["TRCID", "PC_Status"]]

# Merge on TRCID (outer = no data loss)
merged_df = pd.merge(tc_df, pc_df, on="TRCID", how="outer")

# Save to new Excel
merged_df = merged_df.fillna("")
merged_df.to_excel(output_file, index=False)

print("Merged file created successfully!")
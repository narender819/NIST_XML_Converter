"""
Script:
41_append_smiles_to_tc_pc_missing.py

Purpose:
This script enriches the TC/PC missing component dataset by appending CAS Number and SMILES information required for ICAS processing and property recovery activities.

Functionality:
- Reads the consolidated TC/PC missing component report.
- Reads the master SMILES dataset.
- Matches components using TRCID.
- Appends CAS Number and SMILES information to the missing component dataset.
- Identifies components with available SMILES information.
- Identifies components with missing SMILES information.
- Generates separate reports for matched and unmatched SMILES records.
- Creates a consolidated dataset for downstream ICAS processing.

Input:
- Consolidated TC/PC Missing Components Excel report.
- Master SMILES Excel file.

Output:
- TC/PC Missing Components with SMILES Excel report containing:
- ALL_DATA sheet
- SMILES_FOUND sheet
- SMILES_MISSING sheet
"""



import pandas as pd


from config import (
    TC_PC_MERGED_FILE,
    SMILES_MASTER_FILE,
    TC_PC_SMILES_FILE
)

merged_file = TC_PC_MERGED_FILE

smiles_file = SMILES_MASTER_FILE

output_file = TC_PC_SMILES_FILE

# ---------------------------------------------------
# FILE PATHS
# ---------------------------------------------------
# merged_file = r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\2_NIST_TC_PC_Merged.xlsx"
# smiles_file = r"D:\NIST_XML_Converter\output\2025\smiles\1_compounds_smiles_2025.xlsx"

# output_file = r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\3_NIST_TC_PC_SMILES_Merged_sep.xlsx"



# ---------------------------------------------------
# READ FILES
# ---------------------------------------------------
merged_df = pd.read_excel(merged_file)
smiles_df = pd.read_excel(smiles_file, sheet_name="Sheet1")

# ---------------------------------------------------
# CLEAN DATA
# ---------------------------------------------------
merged_df["TRCID"] = merged_df["TRCID"].astype(str).str.strip()
smiles_df["TRCID"] = smiles_df["TRCID"].astype(str).str.strip()

# Keep required columns
smiles_df = smiles_df[["TRCID", "CASRN", "SMILES"]]

# ---------------------------------------------------
# MERGE
# ---------------------------------------------------
final_df = pd.merge(
    merged_df,
    smiles_df,
    on="TRCID",
    how="left"
)

# ---------------------------------------------------
# SPLIT DATA
# ---------------------------------------------------
found_df = final_df[final_df["SMILES"].notna()]
missing_df = final_df[final_df["SMILES"].isna()]

# ---------------------------------------------------
# WRITE TO EXCEL (MULTIPLE SHEETS)
# ---------------------------------------------------
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="ALL_DATA", index=False)
    found_df.to_excel(writer, sheet_name="SMILES_FOUND", index=False)
    missing_df.to_excel(writer, sheet_name="SMILES_MISSING", index=False)

print("SMILES mapping completed with split sheets!")
print(f"Output file saved at: {output_file}")
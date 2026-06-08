"""
Script:
    15_fill_missing_tc_pc_acentricfactor.py

Purpose:
    This script fills missing critical temperature (TC),
    critical pressure (PC), and acentric factor (AF)
    values using external source data.

Functionality:
    - Reads the main thermodynamic property dataset
    - Reads source TC, PC, and AF reference data
    - Normalizes CAS numbers for matching
    - Merges datasets using TRCID and CASRN
    - Fills missing TC, PC, and AF values
    - Preserves existing values when already available
    - Generates fill status remarks
    - Exports the updated thermodynamic property dataset

Input:
    - Main thermodynamic property Excel file
    - Source TC/PC/AF reference Excel file

Output:
    Updated Excel report containing filled TC, PC, and AF values
"""

import pandas as pd

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# PREREQUISITE DIRECTORIES
# ==================================================
PREREQ_DIR = BASE_DIR / "prerequisites"

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR

PROCESSED_DIR = (OUTPUT_DIR/ "processed"/ "full_library")

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================
main_file = (PROCESSED_DIR/ f"10_NIST_Component_KeyThermoProperties_{RUN_YEAR}.xlsx")

source_file = (EXCEL_INPUT_DIR/ "7_TC_PC_AF_extracted.xlsx")

output_file = (PROCESSED_DIR/ f"12_NIST_Component_KeyThermoProperties_{RUN_YEAR}_TCPCAF_UPDATED.xlsx")


# ---------------------------------------------------
# Normalize CAS
# ---------------------------------------------------
def normalize_cas(cas):
    if pd.isna(cas):
        return None
    s = str(cas).strip().replace("-", "")
    if "." in s:
        s = s.split(".")[0]
    return s


# ---------------------------------------------------
# Load Data
# ---------------------------------------------------
print("Loading files...")
main_df = pd.read_excel(main_file)
src_df = pd.read_excel(source_file)

print(f"Main rows: {len(main_df)}, Source rows: {len(src_df)}")


# ---------------------------------------------------
# Normalize CAS
# ---------------------------------------------------
main_df["CAS_norm"] = main_df["CASNO"].apply(normalize_cas)
src_df["CAS_norm"] = src_df["CASRN"].apply(normalize_cas)

# Replace -9999 with NA (for proper filling logic)
main_df["TC (K)"] = main_df["TC (K)"].replace(-9999, pd.NA)
main_df["PC (kPa)"] = main_df["PC (kPa)"].replace(-9999, pd.NA)

# ADD THIS (ACENTRIC normalization)
if "ACENTRIC" in main_df.columns:
    main_df["ACENTRIC"] = main_df["ACENTRIC"].replace(-9999, pd.NA)


# ---------------------------------------------------
# Rename source columns
# ---------------------------------------------------
src_df = src_df.rename(columns={
    "TC": "TC_src",
    "PC": "PC_src",
    "OMEGA": "ACENTRIC_src"   #  NEW
})


# ---------------------------------------------------
# Merge (TRCID + CAS)
# ---------------------------------------------------
print("Merging...")
merged = pd.merge(
    main_df,
    src_df[["TRCID", "CAS_norm", "TC_src", "PC_src", "ACENTRIC_src"]],
    on=["TRCID", "CAS_norm"],
    how="left"
)


# ---------------------------------------------------
# Ensure columns exist
# ---------------------------------------------------
if "TC (K)" not in merged.columns:
    merged["TC (K)"] = pd.NA

if "PC (kPa)" not in merged.columns:
    merged["PC (kPa)"] = pd.NA

if "ACENTRIC" not in merged.columns:
    merged["ACENTRIC"] = pd.NA


# ---------------------------------------------------
# Preserve original values
# ---------------------------------------------------
merged["TC_original"] = merged["TC (K)"]
merged["PC_original"] = merged["PC (kPa)"]
merged["ACENTRIC_original"] = merged["ACENTRIC"]


# ---------------------------------------------------
# Fill logic (only missing)
# ---------------------------------------------------
tc_missing = merged["TC (K)"].isna()
pc_missing = merged["PC (kPa)"].isna()
acentric_missing = merged["ACENTRIC"].isna()

merged.loc[tc_missing & merged["TC_src"].notna(), "TC (K)"] = merged["TC_src"]
merged.loc[pc_missing & merged["PC_src"].notna(), "PC (kPa)"] = merged["PC_src"]
merged.loc[acentric_missing & merged["ACENTRIC_src"].notna(), "ACENTRIC"] = merged["ACENTRIC_src"]


# ---------------------------------------------------
# Remark Logic (UPDATED)
# ---------------------------------------------------
def build_remark(row):

    tc_before = pd.isna(row["TC_original"])
    pc_before = pd.isna(row["PC_original"])
    ac_before = pd.isna(row["ACENTRIC_original"])

    tc_after = pd.notna(row["TC (K)"])
    pc_after = pd.notna(row["PC (kPa)"])
    ac_after = pd.notna(row["ACENTRIC"])

    filled = []

    if tc_before and tc_after:
        filled.append("TC")

    if pc_before and pc_after:
        filled.append("PC")

    if ac_before and ac_after:
        filled.append("ACENTRIC")

    if filled:
        return " & ".join(filled) + " filled"

    if not tc_before and not pc_before and not ac_before:
        return "Already available"

    return "No data"


print("Generating remarks...")
merged["TC_PC_Remark"] = merged.apply(build_remark, axis=1)


# ---------------------------------------------------
# Place Remark column next to TC & PC
# ---------------------------------------------------
cols = list(merged.columns)

tc_index = cols.index("TC (K)")
pc_index = cols.index("PC (kPa)")

cols.remove("TC_PC_Remark")
insert_pos = max(tc_index, pc_index) + 1
cols.insert(insert_pos, "TC_PC_Remark")

merged = merged[cols]


# ---------------------------------------------------
# Clean-up
# ---------------------------------------------------
merged.drop(columns=[
    "TC_src", "PC_src", "ACENTRIC_src",
    "CAS_norm",
    "TC_original", "PC_original", "ACENTRIC_original"
], inplace=True, errors="ignore")


# ---------------------------------------------------
# Convert NA → -9999 (BEFORE EXPORT)
# ---------------------------------------------------
merged["TC (K)"] = merged["TC (K)"].fillna(-9999)
merged["PC (kPa)"] = merged["PC (kPa)"].fillna(-9999)
merged["ACENTRIC"] = merged["ACENTRIC"].fillna(-9999)


# ---------------------------------------------------
# Save Output
# ---------------------------------------------------
print("Writing output...")
merged.to_excel(output_file, index=False)

print("Done")
print(f"Output: {output_file}")
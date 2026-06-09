"""
Script:
   06_append_dippr_family_abbreviations.py

Purpose:
    This script maps DIPPR family and subfamily codes to their
    corresponding abbreviations and appends them to the extracted
    DIPPR family dataset.

Functionality:
    - Reads the extracted DIPPR family Excel data
    - Reads the DIPPR family/subfamily abbreviation mapping file
    - Performs case-insensitive code matching
    - Maps family and subfamily abbreviations
    - Appends abbreviation columns to the main dataset
    - Sorts records by TRCID
    - Generates the final consolidated Excel report

Input:
    - DIPPR family extraction Excel file
    - DIPPR family/subfamily code mapping Excel file

Output:
    Excel report containing DIPPR family and subfamily abbreviations
"""


from pathlib import Path
import pandas as pd
import os

from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITE DIRECTORIES (derived locally)
# ==================================================

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

MAIN_FILE = PROCESSED_DIR / "4_DIPPR_Family_Extraction_All_json.xlsx"

CODES_FILE = EXCEL_INPUT_DIR / "3_DIPPR_Codes_family_subfamily_List.xlsx"

if not MAIN_FILE.exists():
    raise FileNotFoundError(f"Missing input: {MAIN_FILE}")

if not CODES_FILE.exists():
    raise FileNotFoundError(f"Missing input: {CODES_FILE}")

OUTPUT_FILE = PROCESSED_DIR / "5_DIPPR_Family_Extraction_withfamilyAbbrevations.xlsx"

# ------------------------------------------------------
# STEP 1 — Load the two Excel files
# ------------------------------------------------------
main_df = pd.read_excel(MAIN_FILE)
codes_df = pd.read_excel(CODES_FILE)

# Normalize column names to avoid mismatches
codes_df.columns = [c.strip() for c in codes_df.columns]

# Expected column: "DIPPR[family-subfamily]" and "Abbrevations"
if "DIPPR[family-subfamily]" not in codes_df.columns or "Abbrevations" not in codes_df.columns:
    raise ValueError("Columns 'DIPPR[family-subfamily]' and 'Abbrevations' not found in DIPPR codes file!")

# ------------------------------------------------------
# STEP 2 — Normalize codes to uppercase (case-insensitive matching)
# ------------------------------------------------------
codes_df["Code_UP"] = codes_df["DIPPR[family-subfamily]"].astype(str).str.strip().str.upper()
codes_df["Abbrevations"] = codes_df["Abbrevations"].astype(str).str.strip()

# Build lookup dictionary for fast access
lookup_dict = dict(zip(codes_df["Code_UP"], codes_df["Abbrevations"]))

# ------------------------------------------------------
# STEP 3 — Define helper function to map family/subfamily codes
# ------------------------------------------------------
def get_abbreviation(code):
    if pd.isna(code) or str(code).strip() == "":
        return "none"
    key = str(code).strip().upper()
    return lookup_dict.get(key, "none")

# ------------------------------------------------------
# STEP 4 — Apply mapping to main dataset
# ------------------------------------------------------
main_df["DIPPR_Family"] = main_df["DIPPR_Family"].astype(str).str.strip()
main_df["DIPPR_Subfamily"] = main_df["DIPPR_Subfamily"].astype(str).str.strip()

main_df["DIPPR_family_Abbrevations"] = main_df["DIPPR_Family"].apply(get_abbreviation)
main_df["DIPPR_subfamily_Abbrevations"] = main_df["DIPPR_Subfamily"].apply(get_abbreviation)

# ------------------------------------------------------
# STEP 5 — SORT by TRCID before saving
# ------------------------------------------------------
main_df = main_df.sort_values(by="TRCID", ascending=True)

# ------------------------------------------------------
# STEP 6 — Save final merged file
# ------------------------------------------------------
main_df.to_excel(OUTPUT_FILE, index=False)

print(f"Saved output sorted by TRCID {OUTPUT_FILE}")

print("\n========================================")
print(" DIPPR Abbreviation Merge Completed")
print(f" Output File : {OUTPUT_FILE}")
print("========================================\n")

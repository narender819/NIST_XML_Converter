"""
Script:
    08_merge_smiles_with_component_data.py

Purpose:
    Merge SMILES data with component family extraction data
    using TRCID and CASRN.

Functionality:
    - Reads component family extraction data
    - Reads SMILES dataset
    - Matches records using TRCID and CASRN
    - Appends SMILES information to components
    - Handles missing SMILES values
    - Sorts output by TRCID
    - Generates consolidated output file

Input:
    - DIPPR family component Excel file
    - SMILES Excel file

Output:
    - Excel file with merged component and SMILES data
"""

import pandas as pd
import os
from pathlib import Path

from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    OUTPUT_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITE DIRECTORIES (derived locally)
# ==================================================

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

SMILES_DIR = OUTPUT_DIR / "smiles"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

DIPPR_FILE = PROCESSED_DIR / "6_DIPPR_Family_Group_Assigned.xlsx"

SMILES_FILE = SMILES_DIR / f"2_compounds_smiles_{RUN_YEAR}_removedblanks.xlsx"

if not DIPPR_FILE.exists():
    raise FileNotFoundError(f"Missing input: {DIPPR_FILE}")

if not SMILES_FILE.exists():
    raise FileNotFoundError(f"Missing input: {SMILES_FILE}")

OUTPUT_FILE = PROCESSED_DIR / "7_DIPPR_Family_Extraction_All_With_Smiles.xlsx"




def load_excel_file(path, description):
    if not os.path.exists(path):
        print(f"ERROR: {description} file not found: {path}")
        return None
    try:
        return pd.read_excel(path)
    except Exception as e:
        print(f"ERROR reading {description}: {e}")
        return None


def main():
    # Load DIPPR excel (main dataset)
    dippr_df = load_excel_file(DIPPR_FILE, "DIPPR dataset")
    if dippr_df is None:
        return

    # Load SMILES dataset
    smiles_df = load_excel_file(SMILES_FILE, "SMILES dataset")
    if smiles_df is None:
        return

    # Standardize column names (strip spaces, uppercase)
    dippr_df.columns = dippr_df.columns.str.strip()
    smiles_df.columns = smiles_df.columns.str.strip()

    # Clean CAS fields (make string)
    dippr_df["CASRN"] = dippr_df["CASRN"].astype(str).str.strip()
    smiles_df["CASRN"] = smiles_df["CASRN"].astype(str).str.strip()

    # Clean TRCID
    dippr_df["TRCID"] = dippr_df["TRCID"].astype(str).str.strip()
    smiles_df["TRCID"] = smiles_df["TRCID"].astype(str).str.strip()

    # Merge based on TRCID first, CASNO second
    merged = pd.merge(
        dippr_df,
        smiles_df[["TRCID", "CASRN", "SMILES"]],
        how="left",
        left_on=["TRCID", "CASRN"],
        right_on=["TRCID", "CASRN"]
    )

    # Drop CASNO column (duplicate)
    # merged.drop(columns=["CASRN"], inplace=True)
    # Ensure CASRN is preserved as single column
    merged["CASRN"] = merged["CASRN"].astype(str).str.strip()


    # Replace missing SMILES with NA
    merged["SMILES"] = merged["SMILES"].fillna("NA")
     # SORT TRCID IN ASCENDING NUMERIC ORDER
    merged = merged.sort_values(by="TRCID", key=lambda x: x.astype(int), ascending=True)

    # Save output
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            merged.to_excel(writer, sheet_name="all_components", index=False)
        print(f"Output generated: {os.path.abspath(OUTPUT_FILE)}")
    except Exception as e:
        print(f"ERROR writing output Excel: {e}")


if __name__ == "__main__":
    main()
    

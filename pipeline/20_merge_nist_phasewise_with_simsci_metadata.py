"""
Script:
    20_merge_nist_phasewise_with_simsci_metadata.py

Purpose:
    This script merges NIST phasewise thermodynamic workflow data
    with existing SIMSCI component metadata and availability information.

Functionality:
    - Reads NIST phasewise workflow Excel sheets
    - Reads SIMSCI component availability metadata
    - Merges datasets using CAS numbers
    - Adds SIMSCI availability status
    - Preserves all phasewise workflow sheets
    - Generates a consolidated multi-sheet Excel report

Input:
    - NIST phasewise workflow Excel file
    - SIMSCI availability Excel file

Output:
    Multi-sheet Excel report containing merged NIST and SIMSCI metadata
"""


import pandas as pd
import numpy as np

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR

PROCESSED_DIR = (
    OUTPUT_DIR
    / "processed"
    / "full_library"
)

PROCESSED_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================
NIST_EXCEL = (
    PROCESSED_DIR
    / f"14_NIST_Splitsheets_PE1legacy_Hdepart.xlsx"
)

SIMSCI_AVAIL_EXCEL = (
    PROCESSED_DIR
    / f"13_NIST_Components_Available_in_SIMSCI_{RUN_YEAR}.xlsx"
)

OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"15_NIST_Phasewise_With_SIMSCI_Metadata.xlsx"
)

# ---------------------------------------------------
# MERGE LOGIC
# ---------------------------------------------------
def merge_nist_simsci_metadata():

    # Read SIMSCI availability (single sheet)
    simsci_df = pd.read_excel(SIMSCI_AVAIL_EXCEL)

    # Drop technical / tracking columns if present
    simsci_df = simsci_df.drop(
        columns=["S.No", "RowID", "NIST ID"],
        errors="ignore"
    )

    # Normalize CAS as string (safety)
    simsci_df["CASNO"] = simsci_df["CASNO"].astype(str)

    # Read ALL sheets from NIST Excel
    nist_sheets = pd.read_excel(NIST_EXCEL, sheet_name=None)

    output_sheets = {}

    for sheet_name, nist_df in nist_sheets.items():

        nist_df["CASNO"] = nist_df["CASNO"].astype(str)

        # ---- MERGE ON CASNO ----
        merged_df = nist_df.merge(
            simsci_df,
            on="CASNO",
            how="left"
        )

        # ---- ADD AVAILABILITY FLAG ----
        merged_df["Available_in_SIMSCI"] = np.where(
            merged_df["SIMSCI ID"].notna(),
            "Yes",
            "No"
        )

        # ---- OPTIONAL: move flag to end ----
        cols = [c for c in merged_df.columns if c != "Available_in_SIMSCI"]
        cols.append("Available_in_SIMSCI")
        merged_df = merged_df[cols]

        output_sheets[sheet_name] = merged_df

        print(f"Merged SIMSCI metadata into sheet: {sheet_name}")

    # ---- WRITE MULTI-SHEET OUTPUT ----
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        for sheet_name, df in output_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("\nScript-2 completed successfully")
    print("Output written to:", OUTPUT_EXCEL)


if __name__ == "__main__":
    merge_nist_simsci_metadata()

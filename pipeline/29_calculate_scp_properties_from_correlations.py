"""
Script:
    29_calculate_scp_properties_from_correlations.py

Purpose:
    This script calculates SCP thermodynamic properties
    using NIST and SIMSCI Solid Enthalpy correlation equations.

Functionality:
    - Reads SCP workflow and SCP coefficient datasets
    - Extracts NIST and SIMSCI SCP correlation coefficients
    - Evaluates SCP correlation equations
    - Calculates SCP enthalpy properties at:
        * Tmin
        * Tmax
        * Trecon
        * Tnmp
    - Calculates SCP derivative properties
    - Updates SCP workflow sheets with calculated values
    - Preserves existing LCP and ICP workflow sheets
    - Generates cleaned SCP workflow output reports
    - Auto-adjusts Excel column widths

Input:
    - SCP workflow Excel file
    - NIST and SIMSCI SCP coefficient Excel file

Output:
    Updated Excel report containing calculated SCP properties
"""

import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import load_workbook
import time
start = time.time()

from scp_corr_1_18_implementation import CorrelationEngine

def auto_adjust_column_widths(excel_path, max_width=40, padding=2):
    wb = load_workbook(excel_path)

    for ws in wb.worksheets:
        for column_cells in ws.columns:
            max_len = 0
            col_letter = column_cells[0].column_letter

            for cell in column_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            adjusted = min(max_len + padding, max_width)
            ws.column_dimensions[col_letter].width = adjusted

    wb.save(excel_path)

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = (
    BASE_DIR
    / "output"
    / RUN_YEAR
)

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
# INPUT FILES
# ==================================================
MAIN_EXCEL = (
    PROCESSED_DIR
    / f"18_NIST_With_SIMSCI_Trecon_PE1Trecontoexcel.xlsx"
)

KEY_EXCEL = (
    PROCESSED_DIR
    / f"19_NIST_SIMSCI_SCP_Key.xlsx"
)

MAIN_SHEET = "SCP"

NIST_KEY_SHEET = (
    "NIST_SCP_Key"
)

SIMSCI_KEY_SHEET = (
    "SIMSCI_SCP_Key"
)

# ==================================================
# OUTPUT FILE
# ==================================================
OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"20_NIST_SCP_with_calculated_props.xlsx"
)

# =========================================================
# TEMPERATURE → OUTPUT COLUMN MAPPING
# =========================================================
NIST_TEMP_MAP = {
    "SCP_Tnmp (K)": "HS@Tnmp (J/kg-mole)",
    "SCPtmin (K)": "HS@Tmin_NIST (J/kg-mole)",
    "SCPtmax (K)": "HS@Tmax_NIST (J/kg-mole)",
    "SCP_Trecon_NIST (K)": "HS@Trecon_NIST (J/kg-mole)",
}

SIMSCI_TEMP_MAP = {
    "SCP_Trecon_SIMSCI (K)": "HS@Trecon_SIMSCI (J/kg-mole)",
    "SEtmin_simsci": "HS@Tmin_simsci (J/kg-mole)",
    "SEtmax_simsci": "HS@Tmax_simsci (J/kg-mole)",
}

DERIVATIVE_MAP = {
    "SCP_Tnmp (K)": "dHS/dT@Tnmp",
}

# =========================================================
# LOAD DATA
# =========================================================
main_df = pd.read_excel(MAIN_EXCEL, sheet_name=MAIN_SHEET)
nist_key_df = pd.read_excel(KEY_EXCEL, sheet_name=NIST_KEY_SHEET)
simsci_key_df = pd.read_excel(KEY_EXCEL, sheet_name=SIMSCI_KEY_SHEET)

engine = CorrelationEngine()

# =========================================================
# HELPER
# =========================================================
def get_coefficients(row, source="NIST"):
    coeffs = []

    if source == "NIST":
        prefix = "C"
        suffix = ""
    else:  # SIMSCI
        prefix = "C"
        suffix = "_sim"

    i = 1
    while True:
        col = f"{prefix}{i}{suffix}"
        if col not in row:
            break

        val = row.get(col)
        if pd.isna(val):
            break

        coeffs.append(val)
        i += 1

    return coeffs



# =========================================================
# LOOP OVER COMPONENTS
# =========================================================
for idx, row in main_df.iterrows():

    cas = row.get("CASNO")
    if pd.isna(cas):
        continue

    # -----------------------------------------------------
    # NIST SOLID CALCULATIONS
    # -----------------------------------------------------
    nist_row = nist_key_df[nist_key_df["CASNO"] == cas]

    if not nist_row.empty:
        nist_row = nist_row.iloc[0]

        eq_no = int(nist_row.get("SCPEqn", -1))
        coeffs = get_coefficients(nist_row, source="NIST")

        if eq_no in (1, 18) and coeffs:

            for temp_col, output_col in NIST_TEMP_MAP.items():
                T = row.get(temp_col)

                if pd.notna(T):
                    H = engine.evaluate(eq_no, T, coeffs, derivative=False)
                    main_df.at[idx, output_col] = H

            # Derivative if needed
            for temp_col, deriv_col in DERIVATIVE_MAP.items():
                T = row.get(temp_col)
                if pd.notna(T):
                    dH = engine.evaluate(eq_no, T, coeffs, derivative=True)
                    main_df.at[idx, deriv_col] = dH

    # -----------------------------------------------------
    # SIMSCI SOLID CALCULATIONS
    # -----------------------------------------------------
    simsci_row = simsci_key_df[simsci_key_df["CASNO"] == cas]

    if not simsci_row.empty:
        simsci_row = simsci_row.iloc[0]

        eq_no = int(simsci_row.get("SCPEqn_sim", -1))
        coeffs = get_coefficients(simsci_row, source="SIMSCI")

        if eq_no in (1, 18) and coeffs:

            for temp_col, output_col in SIMSCI_TEMP_MAP.items():
                T = row.get(temp_col)

                if pd.notna(T):
                    H = engine.evaluate(eq_no, T, coeffs, derivative=False)
                    main_df.at[idx, output_col] = H


print("SCP property calculations completed.")

# =========================================================
# CLEAN MAIN_DF COLUMNS FOR SCP OUTPUT
# =========================================================
def keep_scp_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Drop columns that clearly belong to key sheets (suffixes _nist/_sim,
    raw coefficient columns, etc.) so SCP sees only its own data + results.
    No hard‑coded full list of names, only patterns.
    """
    cols_to_drop = []

    for col in df.columns:
        name = str(col)

        # 1) Drop all pure coefficient columns like C1..C8, C1_sim..C8_sim
        if name.startswith("C") and (name[1:].isdigit() or name[1:].split("_")[0].isdigit()):
            cols_to_drop.append(col)
            continue

        # 2) Drop obvious key‑sheet columns by suffix or tag
        if any(tag in name for tag in ["_nist", "C*_sim", "TRCID_", "ComponentName_", "Aliases_", "PropUnit"]):
            cols_to_drop.append(col)
            continue

        # 3) Keep calculated & SCP temperature columns (HS@..., dHS/dT..., SCP...)
        # so do nothing for those

    return df.drop(columns=cols_to_drop)


# Clean SCP columns before writing (pattern‑based, not hard‑coded list)
main_df = keep_scp_columns(main_df)

# =========================================================
# WRITE OUTPUT (Simple - preserve LCP/ICP, clean SCP only)
# =========================================================
with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    lcp_df = pd.read_excel(MAIN_EXCEL, sheet_name='LCP')
    icp_df = pd.read_excel(MAIN_EXCEL, sheet_name='ICP')
    lcp_df.to_excel(writer, sheet_name='LCP', index=False)
    icp_df.to_excel(writer, sheet_name='ICP', index=False)
    main_df.to_excel(writer, sheet_name=MAIN_SHEET, index=False)
auto_adjust_column_widths(OUTPUT_EXCEL)
print("Done. Clean SCP + original LCP/ICP preserved.")
print("Output:", OUTPUT_EXCEL)


print(f"Runtime: {(time.time()-start):.2f} sec")
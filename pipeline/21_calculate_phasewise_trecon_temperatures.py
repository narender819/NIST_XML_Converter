"""
Script:
    21_calculate_phasewise_trecon_temperatures.py

Purpose:
    This script calculates phasewise Trecon temperatures for
    LCP, ICP, and SCP workflows using overlapping NIST and
    SIMSCI correlation temperature ranges.

Functionality:
    - Reads phasewise NIST-SIMSCI workflow sheets
    - Calculates Trecon temperatures for:
        * LCP
        * ICP
        * SCP
    - Handles overlap and non-overlap temperature ranges
    - Applies small and large negative delta handling logic
    - Generates Trecon calculation remarks
    - Writes updated phasewise workflow sheets
    - Auto-adjusts Excel column widths

Input:
    Phasewise NIST-SIMSCI workflow Excel file

Output:
    Excel report containing calculated phasewise Trecon temperatures
"""

import pandas as pd

from openpyxl import load_workbook

from utils import auto_adjust_column_widths

from pathlib import Path

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

NEG_TOL = 1.0  # K

def compute_trecon(
    df,
    prop_name,
    exists_col,
    nist_tmin_col,
    nist_tmax_col,
    simsci_tmin_col,
    simsci_tmax_col
):

    trecon_sim_col  = f"{prop_name}_Trecon_SIMSCI (K)"
    trecon_nist_col = f"{prop_name}_Trecon_NIST (K)"
    remark_col      = f"{prop_name}_Trecon_Remark"

    df[trecon_sim_col]  = pd.NA
    df[trecon_nist_col] = pd.NA
    df[remark_col]      = ""

    for idx, row in df.iterrows():

        if row.get(exists_col) != "Yes":
            continue

        if row.get("Available_in_SIMSCI") != "Yes":
            df.at[idx, remark_col] = "SIMSCI_NOT_AVAILABLE"
            continue

        tmin_nist = row.get(nist_tmin_col)
        tmax_nist = row.get(nist_tmax_col)
        tmin_sim  = row.get(simsci_tmin_col)
        tmax_sim  = row.get(simsci_tmax_col)

        if pd.isna(tmin_nist) or pd.isna(tmax_nist) or pd.isna(tmin_sim) or pd.isna(tmax_sim):
            df.at[idx, remark_col] = "MISSING_LIMITS"
            continue

        tmin  = max(tmin_nist, tmin_sim)
        tmax  = min(tmax_nist, tmax_sim)
        delta = tmax - tmin

        # --------------------
        # Δ ≥ 0 : overlap
        # --------------------
        if delta >= 0:
            TT = tmin + delta / 3.3

            df.at[idx, trecon_sim_col]  = TT
            df.at[idx, trecon_nist_col] = TT
            df.at[idx, remark_col]      = "DELTA>=0_OVERLAP"
            continue

        abs_delta = abs(delta)

        # --------------------
        # Small negative Δ
        # --------------------
        if abs_delta <= NEG_TOL:

            if (tmax_nist - tmin_sim) < 0:
                df.at[idx, trecon_sim_col]  = tmin_sim
                df.at[idx, trecon_nist_col] = tmax_nist
                df.at[idx, remark_col]      = "SMALL_NEG_NIST_LEFT"

            elif (tmax_sim - tmin_nist) < 0:
                df.at[idx, trecon_sim_col]  = tmax_sim
                df.at[idx, trecon_nist_col] = tmin_nist
                df.at[idx, remark_col]      = "SMALL_NEG_SIMSCI_LEFT"

            else:
                df.at[idx, remark_col] = "SMALL_NEG_LOGIC_ERROR"

            continue

        # --------------------
        # Large negative Δ
        # --------------------
        if abs_delta > NEG_TOL:

            if (tmax_nist - tmin_sim) < 0:
                df.at[idx, trecon_sim_col]  = tmax_nist
                df.at[idx, trecon_nist_col] = tmax_nist
                df.at[idx, remark_col]      = "LARGE_NEG_NO_EXTRAP_NIST_MAX"

            elif (tmax_sim - tmin_nist) < 0:
                df.at[idx, trecon_sim_col]  = tmin_nist
                df.at[idx, trecon_nist_col] = tmin_nist
                df.at[idx, remark_col]      = "LARGE_NEG_NO_EXTRAP_NIST_MIN"

            else:
                df.at[idx, remark_col] = "LARGE_NEG_LOGIC_ERROR"

    return df



def main():
    EXCEL_PATH = PROCESSED_DIR / f"15_NIST_Phasewise_With_SIMSCI_Metadata.xlsx"

    OUTPUT_PATH = PROCESSED_DIR / f"16_NIST_With_SIMSCI_Trecon.xlsx"

    #  Validation (recommended)
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Missing input: {EXCEL_PATH}")

    print("Input:", EXCEL_PATH)
    print("Output:", OUTPUT_PATH)

    # your processing logic continues here...

    sheets = pd.read_excel(EXCEL_PATH, sheet_name=None)

    # ---------------- LCP ----------------
    sheets["LCP"] = compute_trecon(
        sheets["LCP"],
        prop_name="LCP",
        exists_col="LCP_Exists",
        nist_tmin_col="LCPtmin (K)",
        nist_tmax_col="LCPtmax (K)",
        simsci_tmin_col="LEtmin_simsci",
        simsci_tmax_col="LEtmax_simsci",
    )

    # ---------------- ICP ----------------
    sheets["ICP"] = compute_trecon(
        sheets["ICP"],
        prop_name="ICP",
        exists_col="ICP_Exists",
        nist_tmin_col="ICPtmin (K)",
        nist_tmax_col="ICPtmax (K)",
        simsci_tmin_col="IEtmin_simsci",
        simsci_tmax_col="IEtmax_simsci",
    )

    # ---------------- SCP ----------------
    sheets["SCP"] = compute_trecon(
        sheets["SCP"],
        prop_name="SCP",
        exists_col="SCP_Exists",
        nist_tmin_col="SCPtmin (K)",
        nist_tmax_col="SCPtmax (K)",
        simsci_tmin_col="SEtmin_simsci",
        simsci_tmax_col="SEtmax_simsci",
    )

    # Write output
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    auto_adjust_column_widths(OUTPUT_PATH)
    print("Trecon calculation completed for LCP / ICP / SCP")



if __name__ == "__main__":
    main()



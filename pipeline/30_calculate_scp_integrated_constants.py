"""
Script:
    30_calculate_scp_integrated_constants.py

Purpose:
    This script calculates SCP integrated constants (IC)
    using NIST and SIMSCI solid enthalpy reconciliation logic.

Functionality:
    - Reads SCP thermodynamic workflow sheets
    - Calculates SCP integrated constants using:
        * SIMSCI protocol logic
        * NIST-SIMSCI reconciliation logic
    - Handles overlap and non-overlap temperature ranges
    - Calculates Delta_C1 and adjusted C1 values
    - Calculates ICSolid values
    - Applies heat of fusion corrections
    - Handles extrapolation and missing temperature limits
    - Updates SCP workflow sheets with calculated IC values
    - Preserves existing LCP and ICP sheets
    - Auto-adjusts Excel column widths

Input:
    SCP thermodynamic workflow Excel file

Output:
    Updated Excel report containing calculated SCP integrated constants
"""


import pandas as pd
import numpy as np
from openpyxl import load_workbook

import time
start = time.time()

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
# INPUT / OUTPUT FILES
# ==================================================
INPUT_EXCEL = (
    PROCESSED_DIR
    / f"20_NIST_SCP_with_calculated_props.xlsx"
)

OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"21_NIST_ICcalc_SCP.xlsx"
)

# ---------------------------------------------------
# CONSTANTS
# ---------------------------------------------------
NEG_TOL = 1.0
sentinel = -9999

# ---------------------------------------------------
# HELPERS
# ---------------------------------------------------
def linear_extrapolate(H_ref, T_ref, dHdT_ref, T):
    return H_ref + (T - T_ref) * dHdT_ref


# ---------------------------------------------------
# SIMSCI PROTOCOL – SCP
# ---------------------------------------------------
def calc_scp_simsci(row):
    """
    SIMSCI Solid Enthalpy IC calculation (no reconciliation).
    """

    if row.get("SCP_Exists") != "Yes":
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICSolid (J/kg-mole)": pd.NA,
            "ICSolid_Method": "SIMSCI",
            "ICSolid_Status": "NO_SCP"
        })

    # -------------------------------------------------
    # Step 1: Define tnmp
    # -------------------------------------------------
    if pd.notna(row.get("NMP (K)")) and row["NMP (K)"] != sentinel:
        tnmp = row["NMP (K)"]
    else:
        tmin = row.get("SCPtmin (K)")
        tmax = row.get("SCPtmax (K)")
        if pd.isna(tmin) or pd.isna(tmax):
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICSolid (J/kg-mole)": pd.NA,
                "ICSolid_Method": "SIMSCI",
                "ICSolid_Status": "SCP_LIMITS_MISSING"
            })
        tnmp = tmin + 0.95 * (tmax - tmin)

    # -------------------------------------------------
    # Step 2: Heat of fusion at NMP
    # -------------------------------------------------
    if pd.notna(row.get("HFUSIONNMP (kJ/kg-mol)")):
        hhfusion = row["HFUSIONNMP (kJ/kg-mol)"] * 1000.0
    else:
        hhfusion = 0.0

    # -------------------------------------------------
    # Step 3: HL at tnmp (extrapolate if required)
    # -------------------------------------------------
    HL = 0.0
    if row.get("LCP_Exists") == "Yes":
        tmin = row.get("LCPtmin (K)")
        tmax = row.get("LCPtmax (K)")

        if pd.notna(tmin) and pd.notna(tmax):
            if tnmp < tmin:
                HL = (
                    row.get("HL@Tmin (J/kg-mole)") +
                    (tnmp - tmin) * row.get("dHL/dT@Tmin", 0.0)
                )
            elif tnmp > tmax:
                HL = (
                    row.get("HL@Tmax (J/kg-mole)") +
                    (tnmp - tmax) * row.get("dHL/dT@Tmax", 0.0)
                )
            else:
                HL = row.get("HL@Tnmp (J/kg-mole)", 0.0)

    # -------------------------------------------------
    # Step 4: Solve ICSolid
    #
    # Hsolid = HL - HFUSION
    # ICSolid + HSolid(Tnmp) = HL - HFUSION
    # -------------------------------------------------
    Hsolid = HL - hhfusion
    HS = row.get("HS@Tnmp (J/kg-mole)", 0.0)

    ICSolid = Hsolid - HS

    return pd.Series({
        "Delta_C1 (J/kg-mole)": ICSolid,
        "C1_Adjusted (J/kg-mole)": ICSolid/1000,
        "ICSolid (J/kg-mole)": ICSolid/1000,
        "ICSolid_Method": "SIMSCI-Protocol",
        "ICSolid_Status": "OK"
    })


# ---------------------------------------------------
# SCP ENTHALPY RECONCILIATION (NIST ↔ SIMSCI)
# ---------------------------------------------------
def calc_scp_reconciliation(row):
    """
    SCP reconciliation using SAME Δ logic as LCP/ICP.
    """

    if row.get("SCP_Exists") != "Yes":
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICSolid (J/kg-mole)": pd.NA,
            "ICSolid_Method": "Recon",
            "ICSolid_Status": "NO_SCP"
        })

    # ---------------------------------
    # Solid enthalpies from PropEval
    # ---------------------------------
    HS_Trecon_NIST   = row.get("HS@Trecon_NIST (J/kg-mole)")
    HS_Trecon_SIMSCI = row.get("HS@Trecon_SIMSCI (J/kg-mole)")

    HS_Tmin_NIST   = row.get("HS@Tmin_NIST (J/kg-mole)")
    HS_Tmax_NIST   = row.get("HS@Tmax_NIST (J/kg-mole)")
    HS_Tmin_SIMSCI = row.get("HS@Tmin_simsci (J/kg-mole)")
    HS_Tmax_SIMSCI = row.get("HS@Tmax_simsci (J/kg-mole)")

    C1_NIST = row.get("C1_NIST (J/kg-mole)", 0.0)

    # ---------------------------------
    # Temperature limits
    # ---------------------------------
    tmin_nist   = row.get("SCPtmin (K)")
    tmax_nist   = row.get("SCPtmax (K)")
    tmin_simsci = row.get("SEtmin_simsci")
    tmax_simsci = row.get("SEtmax_simsci")

    if any(pd.isna(v) for v in [tmin_nist, tmax_nist, tmin_simsci, tmax_simsci]):
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICSolid (J/kg-mole)": pd.NA,
            "ICSolid_Method": "Recon",
            "ICSolid_Status": "LIMITS_MISSING"
        })

    # ---------------------------------
    # Δ logic
    # ---------------------------------
    tmin = max(tmin_nist, tmin_simsci)
    tmax = min(tmax_nist, tmax_simsci)
    Delta = tmax - tmin

    if Delta >= 0:
        H_nist = HS_Trecon_NIST
        H_sim  = HS_Trecon_SIMSCI
        recon_case = "Delta>=0"

    elif abs(Delta) <= NEG_TOL:
        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HS_Tmax_NIST
            H_sim  = HS_Tmin_SIMSCI
            recon_case = "SmallNeg:NISTmax_SIMSCImin"
        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HS_Tmin_NIST
            H_sim  = HS_Tmax_SIMSCI
            recon_case = "SmallNeg:SIMSCImax_NISTmin"
        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICSolid (J/kg-mole)": pd.NA,
                "ICSolid_Method": "Recon",
                "ICSolid_Status": "TEMP_LOGIC_ERROR"
            })
    else:
        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HS_Tmax_NIST
            H_sim  = HS_Tmax_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmax"
        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HS_Tmin_NIST
            H_sim  = HS_Tmin_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmin"
        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICSolid (J/kg-mole)": pd.NA,
                "ICSolid_Method": "Recon",
                "ICSolid_Status": "TEMP_LOGIC_ERROR"
            })

    if pd.isna(H_nist) or pd.isna(H_sim):
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICSolid (J/kg-mole)": pd.NA,
            "ICSolid_Method": "Recon",
            "ICSolid_Status": "ENTHALPY_MISSING"
        })

    Delta_C1 = H_sim - H_nist
    C1_Adjusted = C1_NIST + Delta_C1

    return pd.Series({
        "Delta_C1 (J/kg-mole)": Delta_C1,
        "C1_Adjusted (J/kg-mole)": C1_Adjusted/1000,
        "ICSolid (J/kg-mole)": C1_Adjusted/1000,
        "ICSolid_Method": f"SIMSCI-Recon[{recon_case}]",
        "ICSolid_Status": "OK"
    })


# ---------------------------------------------------
# DRIVER
# ---------------------------------------------------
# def calc_scp(row):
#     if row["Available_in_SIMSCI"] == "No":
#         return calc_scp_simsci(row)
#     else:
#         return calc_scp_reconciliation(row)

def calc_scp(row):

    # If not available in SIMSCI → direct SIMSCI protocol
    if row.get("Available_in_SIMSCI") == "No":
        return calc_scp_simsci(row)

    # If available in SIMSCI and SCP exists
    if row.get("Available_in_SIMSCI") == "Yes" and row.get("SCP_Exists") == "Yes":

        tmin_simsci = row.get("SEtmin_simsci")
        tmax_simsci = row.get("SEtmax_simsci")

        # If SIMSCI temperature limits missing → fallback to SIMSCI protocol
        if any(pd.isna(v) for v in [tmin_simsci, tmax_simsci]):
            return calc_scp_simsci(row)

    # Otherwise → reconciliation
    return calc_scp_reconciliation(row)


def main():

    excel = pd.ExcelFile(INPUT_EXCEL)

    lcp_df = pd.read_excel(excel, sheet_name="LCP")
    icp_df = pd.read_excel(excel, sheet_name="ICP")
    scp_df = pd.read_excel(excel, sheet_name="SCP")

    results = scp_df.apply(calc_scp, axis=1)
    scp_out = pd.concat([scp_df, results], axis=1)

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        lcp_df.to_excel(writer, sheet_name="LCP", index=False)
        icp_df.to_excel(writer, sheet_name="ICP", index=False)
        scp_out.to_excel(writer, sheet_name="SCP", index=False)
    
    auto_adjust_column_widths(OUTPUT_EXCEL)

    print("SCP IC calculation completed successfully")


if __name__ == "__main__":
    main()

print(f"\nRuntime: {(time.time()-start):.2f} sec")
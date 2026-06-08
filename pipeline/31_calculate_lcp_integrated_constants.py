"""
Script:
    31_calculate_lcp_integrated_constants.py

Purpose:
    This script calculates LCP integrated constants (IC)
    using NIST and SIMSCI liquid enthalpy reconciliation logic.

Functionality:
    - Reads LCP thermodynamic workflow sheets
    - Calculates liquid integrated constants using:
        * SIMSCI protocol logic
        * NIST-SIMSCI reconciliation logic
    - Handles overlap and non-overlap temperature ranges
    - Calculates Delta_C1 and adjusted C1 values
    - Calculates ICliq values
    - Applies extrapolation and Tc-based calculations
    - Handles missing temperature limits and enthalpy data
    - Updates LCP workflow sheets with calculated IC values
    - Preserves existing ICP and SCP sheets
    - Auto-adjusts Excel column widths

Input:
    LCP thermodynamic workflow Excel file

Output:
    Updated Excel report containing calculated LCP integrated constants
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
    / f"21_NIST_ICcalc_SCP.xlsx"
)

OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"22_NIST_ICCalc_SCP_LCP.xlsx"
)

# ---------------------------------------------------
# CONSTANTS
# ---------------------------------------------------
TREF = 273.15  # 0 °C

# ---------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------
def extrapolate(H_ref, T_ref, dHdT_ref, T):
    """
    Linear extrapolation:
    H(T) = H(ref) + (T - Tref) * dH/dT(ref)
    """
    return H_ref + (T - T_ref) * dHdT_ref


def calc_lcp_two_step(row):
    """
    Tmax -> Tc -> Tref extrapolation (SIMSCI)
    """
    # Step 1: Tmax -> Tc
    HL_TC = extrapolate(
        row["HL@Tmax (J/kg-mole)"],
        row["LCPtmax (K)"],
        row["dHL/dT@Tmax"],
        row["TC (K)"]
    )

    # Step 2: Tc -> Tref using NBP derivative
    return extrapolate(
        HL_TC,
        row["TC (K)"],
        row["dHL/dT@NBP"],
        TREF
    )

# ---------------------------------------------------
# SIMSCI PROTOCOL – LCP
# ---------------------------------------------------
def calc_lcp_simsci(row):
    if row.get("LCP_Exists") != "Yes":
        return pd.Series({
            "ICliq (J/kg-mole)": np.nan,
            "ICliq_Method": "SIMSCI",
            "ICliq_Status": "NO_LCP"
        })

    Tmin = row.get("LCPtmin (K)")
    Tmax = row.get("LCPtmax (K)")

    if pd.isna(Tmin) or pd.isna(Tmax):
        return pd.Series({
            "ICliq (J/kg-mole)": np.nan,
            "ICliq_Method": "SIMSCI",
            "ICliq_Status": "LIMITS_MISSING"
        })

    # 1. Direct SIMSCI if 0C is valid
    if (
        pd.notna(row.get("HL@0C (J/kg-mole)"))
        and Tmin <= TREF <= Tmax
    ):
        return pd.Series({
            "ICliq (J/kg-mole)": -row["HL@0C (J/kg-mole)"],
            "ICliq_Method": "SIMSCI-Direct",
            "ICliq_Status": "OK"
        })

    # 2. T < Tmin extrapolation
    if TREF < Tmin:
        if pd.isna(row.get("HL@Tmin (J/kg-mole)")) or pd.isna(row.get("dHL/dT@Tmin")):
            return pd.Series({
                "ICliq (J/kg-mole)": np.nan,
                "ICliq_Method": "SIMSCI-Extrapolated",
                "ICliq_Status": "DATA_MISSING"
            })

        H = extrapolate(
            row["HL@Tmin (J/kg-mole)"],
            Tmin,
            row["dHL/dT@Tmin"],
            TREF
        )
        return pd.Series({
            "ICliq (J/kg-mole)": -H,
            "ICliq_Method": "SIMSCI-Extrapolated(T<Tmin)",
            "ICliq_Status": "OK"
        })

    # 3. Two-step Tc logic
    if TREF > row.get("TC (K)", np.inf):
        req = [
            "HL@Tmax (J/kg-mole)",
            "dHL/dT@Tmax",
            "dHL/dT@NBP",
            "TC (K)"
        ]
        if any(pd.isna(row.get(x)) for x in req):
            return pd.Series({
                "ICliq (J/kg-mole)": np.nan,
                "ICliq_Method": "SIMSCI-TwoStep",
                "ICliq_Status": "DATA_MISSING"
            })

        H = calc_lcp_two_step(row)
        return pd.Series({
            "ICliq (J/kg-mole)": -H,
            "ICliq_Method": "SIMSCI-TwoStep(Tc)",
            "ICliq_Status": "OK"
        })

    return pd.Series({
        "ICliq (J/kg-mole)": np.nan,
        "ICliq_Method": "SIMSCI",
        "ICliq_Status": "FAILED"
    })


#############################################
def calc_lcp_reconciliation(row, NegTol=1.0):
    """
    Full SIMSCI-compliant LCP enthalpy reconciliation using
    PropEval-evaluated enthalpies at Trecon / Tmin / Tmax.
    """

    # -------------------------------------------------
    # Preconditions
    # -------------------------------------------------
    if row.get("LCP_Exists") != "Yes":
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICliq (J/kg-mole)": pd.NA,
            "ICliq_Method": "Recon",
            "ICliq_Status": "NO_LCP"
        })

    # -------------------------------------------------
    # Required enthalpy inputs (from PropEval)
    # -------------------------------------------------
    HL_Trecon_NIST   = row.get("HL@Trecon_NIST (J/kg-mole)")
    HL_Trecon_SIMSCI = row.get("HL@Trecon_SIMSCI (J/kg-mole)")

    HL_Tmin_NIST   = row.get("HL@Tmin_NIST (J/kg-mole)")
    HL_Tmax_NIST   = row.get("HL@Tmax_NIST (J/kg-mole)")
    HL_Tmin_SIMSCI = row.get("HL@Tmin_SIMSCI (J/kg-mole)")
    HL_Tmax_SIMSCI = row.get("HL@Tmax_SIMSCI (J/kg-mole)")

    HL_0C = row.get("HL@0C (J/kg-mole)")
    C1_NIST = row.get("C1_NIST (J/kg-mole)", 0.0)

    # Temperature limits
    tmin_nist   = row.get("LCPtmin (K)")
    tmax_nist   = row.get("LCPtmax (K)")
    tmin_simsci = row.get("LEtmin_simsci")
    tmax_simsci = row.get("LEtmax_simsci")

    # -------------------------------------------------
    # Validation
    # -------------------------------------------------
    required = [
        HL_0C, C1_NIST,
        tmin_nist, tmax_nist,
        tmin_simsci, tmax_simsci
    ]

    if any(pd.isna(v) for v in required):
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICliq (J/kg-mole)": pd.NA,
            "ICliq_Method": "Recon",
            "ICliq_Status": "DATA_MISSING"
        })

    # -------------------------------------------------
    # Delta logic
    # -------------------------------------------------
    tmin = max(tmin_nist, tmin_simsci)
    tmax = min(tmax_nist, tmax_simsci)
    Delta = tmax - tmin

    # =================================================
    # Case 1: Δ ≥ 0  (overlap → use Trecon)
    # =================================================
    if Delta >= 0:
        if pd.isna(HL_Trecon_NIST) or pd.isna(HL_Trecon_SIMSCI):
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICliq (J/kg-mole)": pd.NA,
                "ICliq_Method": "Recon",
                "ICliq_Status": "TRECON_DATA_MISSING"
            })

        H_nist = HL_Trecon_NIST
        H_sim  = HL_Trecon_SIMSCI
        recon_case = "Delta>=0"

    # =================================================
    # Case 2: Small negative Δ
    # =================================================
    elif abs(Delta) <= NegTol:

        # (a) TmaxNIST < TminSIMSCI
        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HL_Tmax_NIST
            H_sim  = HL_Tmin_SIMSCI
            recon_case = "SmallNeg:NISTmax_SIMSCImin"

        # (b) TmaxSIMSCI < TminNIST
        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HL_Tmin_NIST
            H_sim  = HL_Tmax_SIMSCI
            recon_case = "SmallNeg:SIMSCImax_NISTmin"

        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICliq (J/kg-mole)": pd.NA,
                "ICliq_Method": "Recon",
                "ICliq_Status": "TEMP_LOGIC_ERROR"
            })

        if pd.isna(H_nist) or pd.isna(H_sim):
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICliq (J/kg-mole)": pd.NA,
                "ICliq_Method": "Recon",
                "ICliq_Status": "BOUNDARY_H_MISSING"
            })

    # =================================================
    # Case 3: Large negative Δ
    # =================================================
    else:
        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HL_Tmax_NIST
            H_sim  = HL_Tmax_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmax"

        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HL_Tmin_NIST
            H_sim  = HL_Tmin_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmin"

        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICliq (J/kg-mole)": pd.NA,
                "ICliq_Method": "Recon",
                "ICliq_Status": "TEMP_LOGIC_ERROR"
            })

    # -------------------------------------------------
    # ΔC1 and ICliq calculation (FINAL RULE)
    # -------------------------------------------------
    Delta_C1 = H_sim - H_nist
    C1_Adjusted = C1_NIST + Delta_C1
    ICliq=C1_Adjusted 
    # ICliq = - (HL_0C + Delta_C1)

    return pd.Series({
        "Delta_C1 (J/kg-mole)": Delta_C1,
        "C1_Adjusted (J/kg-mole)": C1_Adjusted,
        "ICliq (J/kg-mole)": ICliq,
        "ICliq_Method": f"SIMSCI-Recon[{recon_case}]",
        "ICliq_Status": "OK"
    })




# ---------------------------------------------------
# MAIN DRIVER
# ---------------------------------------------------
def calc_lcp(row):
    # STEP 1: Protocol selection
    if row["Available_in_SIMSCI"] == "No":
        return calc_lcp_simsci(row)
    else:
        return calc_lcp_reconciliation(row)


# def main():
#     df = pd.read_excel(INPUT_EXCEL)

#     results = df.apply(calc_lcp, axis=1)

#     out = pd.concat(
#         [df[["TRCID", "CASNO", "ComponentName"]], results],
#         axis=1
#     )

#     out.to_excel(OUTPUT_EXCEL, index=False)
#     print("LCP calculation completed successfully")

def main():
    xls = pd.ExcelFile(INPUT_EXCEL)

    # Read sheets
    lcp_df = pd.read_excel(xls, sheet_name="LCP")
    icp_df = pd.read_excel(xls, sheet_name="ICP")
    scp_df = pd.read_excel(xls, sheet_name="SCP")

    # ---- LCP calculation ONLY ----
    lcp_results = lcp_df.apply(calc_lcp, axis=1)

    # Append results to LCP sheet
    lcp_out = pd.concat([lcp_df, lcp_results], axis=1)

    # ---- Write new Excel ----
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        lcp_out.to_excel(writer, sheet_name="LCP", index=False)
        icp_df.to_excel(writer, sheet_name="ICP", index=False)
        scp_df.to_excel(writer, sheet_name="SCP", index=False)
    
    auto_adjust_column_widths(OUTPUT_EXCEL)

    print("LCP IC calculation completed; ICP & SCP preserved")



if __name__ == "__main__":
    main()

print(f"\nRuntime: {(time.time()-start):.2f} sec")
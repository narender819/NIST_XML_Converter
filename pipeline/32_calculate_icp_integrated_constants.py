"""
Script:
    32_calculate_icp_integrated_constants.py

Purpose:
    This script calculates ICP integrated constants (IC)
    using NIST and SIMSCI ideal gas enthalpy reconciliation logic.

Functionality:
    - Reads ICP thermodynamic workflow sheets
    - Calculates ideal gas integrated constants using:
        * SIMSCI protocol logic
        * NIST-SIMSCI reconciliation logic
    - Handles overlap and non-overlap temperature ranges
    - Calculates Delta_C1 and adjusted C1 values
    - Calculates ICigas values
    - Applies vaporization enthalpy and HDeparture corrections
    - Handles extrapolation and missing temperature limits
    - Supports equation-specific adjustment logic
    - Updates ICP workflow sheets with calculated IC values
    - Preserves existing LCP and SCP sheets
    - Auto-adjusts Excel column widths

Input:
    ICP thermodynamic workflow Excel file

Output:
    Updated Excel report containing calculated ICP integrated constants
"""


import pandas as pd
import numpy as np
from openpyxl import load_workbook
from pathlib import Path

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
    / f"22_NIST_ICCalc_SCP_LCP.xlsx"
)

OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"23_NIST_ICCalc_SCP_LCP_ICP.xlsx"
)

# ---------------------------------------------------
# CONSTANTS
# ---------------------------------------------------
NEG_TOL = 1.0
sentinel = -9999 
# ---------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------
def linear_extrapolate(H_ref, T_ref, dHdT_ref, T):
    return H_ref + (T - T_ref) * dHdT_ref

def calc_icp_simsci(row):
    """
    SIMSCI Ideal Gas Enthalpy protocol with extrapolation.
    No reconciliation is applied here.
    """

    if row.get("ICP_Exists") != "Yes":
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICigas (J/kg-mole)": pd.NA,
            "ICigas_Method": "SIMSCI",
            "ICigas_Status": "NO_ICP"
        })

    # -------------------------------------------------
    # Step 1: Define tnbp
    # -------------------------------------------------
    if pd.notna(row.get("NBP (K)")) and row["NBP (K)"] != sentinel:
        tnbp = row["NBP (K)"]
    else:
        tmin = row.get("ICPtmin (K)")
        tmax = row.get("ICPtmax (K)")
        if pd.isna(tmin) or pd.isna(tmax):
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICigas (J/kg-mole)": pd.NA,
                "ICigas_Method": "SIMSCI",
                "ICigas_Status": "ICP_LIMITS_MISSING"
            })
        tnbp = tmin + 0.05 * (tmax - tmin)

    # ---------------------------------
    # HVAP at tnbp (SIMSCI protocol – NO extrapolation)
    # ---------------------------------
    HVAP = 0.0

    if row.get("HVAP_Exists") == "Yes":

        tmin = row.get("HVAPtmin (K)")
        tmax = row.get("HVAPtmax (K)")

        if pd.notna(tmin) and pd.notna(tmax):

            # CASE B: within range
            if tmin <= tnbp <= tmax:
                HVAP = row.get("HVAP@Tnbp (J/kg-mole)", 0.0)

            # CASE A/C: outside range → midpoint
            else:
                HVAP = row.get(
                    "HVAP@Mid (J/kg-mole)", 
                    0.0
                )


    # -------------------------------------------------
    # Step 3: HL at tnbp (extrapolate if needed)
    # -------------------------------------------------
    HL = 0.0
    if row.get("LCP_Exists") == "Yes":
        tmin = row.get("LCPtmin (K)")
        tmax = row.get("LCPtmax (K)")

        if pd.notna(tmin) and pd.notna(tmax):
            if tnbp < tmin:
                HL_Tmin = row.get("HL@Tmin (J/kg-mole)")
                dHLdTmin   = row.get("dHL/dT@Tmin")

                HL_Tmin = 0.0 if pd.isna(HL_Tmin) else HL_Tmin
                dHLdTmin   = 0.0 if pd.isna(dHLdTmin) else dHLdTmin

                HL = HL_Tmin + (tnbp - tmin) * dHLdTmin
                # HL = (
                #     row.get("HL@Tmin (J/kg-mole)" or 0.0) +
                #     (tnbp - tmin) * row.get("dHL/dT@Tmin", 0.0)
                # )
            elif tnbp > tmax:
                HL_Tmax = row.get("HL@Tmax (J/kg-mole)")
                dHLdTmax   = row.get("dHL/dT@Tmax")
                HL_Tmax = 0.0 if pd.isna(HL_Tmax) else HL_Tmax
                dHLdTmax   = 0.0 if pd.isna(dHLdTmax) else dHLdTmax

                HL = HL_Tmax + (tnbp - tmax) * dHLdTmax
                
                # HL = (
                #     row.get("HL@Tmax (J/kg-mole)" or 0.0) +
                #     (tnbp - tmax) * row.get("dHL/dT@Tmax", 0.0)
                # )
            else:
                HL = row.get("HL@Tnbp (J/kg-mole)", 0.0)

    # -------------------------------------------------
    # Step 4: Hdeparture
    # -------------------------------------------------
    Hdep = row.get("HDeparture (J/kg-mole)", 0.0)

    # -------------------------------------------------
    # Step 5: Solve ICigas
    #
    # Hideal = HL + HVAP - Hdeparture
    # ICigas + HIG(Tnbp) = Hideal
    # -------------------------------------------------
    Hideal = HL + HVAP - Hdep
    HIG = row.get("HIG@Tnbp (J/kg-mole)", 0.0)

    ICigas = Hideal - HIG

    return pd.Series({
        "Delta_C1 (J/kg-mole)": ICigas,
        "C1_Adjusted (J/kg-mole)": ICigas,
        "ICigas (J/kg-mole)": ICigas,
        "ICigas_Method": "SIMSCI-Protocol",
        "ICigas_Status": "OK"
    })


# ---------------------------------------------------
# ICP ENTHALPY RECONCILIATION (NIST ↔ SIMSCI)
# ---------------------------------------------------
def calc_icp_reconciliation(row):
    """
    ICP reconciliation using the SAME Δ logic as LCP,
    but with ICP enthalpy values.
    """

    if row.get("ICP_Exists") != "Yes":
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICigas (J/kg-mole)": pd.NA,
            "ICigas_Method": "Recon",
            "ICigas_Status": "NO_ICP"
        })

    # ---------------------------------
    # Required enthalpies (PropEval)
    # ---------------------------------
    HIG_Trecon_NIST   = row.get("HIG@Trecon_NIST (J/kg-mole)")
    HIG_Trecon_SIMSCI = row.get("HIG@Trecon_SIMSCI (J/kg-mole)")

    HIG_Tmin_NIST   = row.get("HIG@Tmin_NIST (J/kg-mole)")
    HIG_Tmax_NIST   = row.get("HIG@Tmax_NIST (J/kg-mole)")
    HIG_Tmin_SIMSCI = row.get("HIG@Tmin_SIMSCI (J/kg-mole)")
    HIG_Tmax_SIMSCI = row.get("HIG@Tmax_SIMSCI (J/kg-mole)")

    C1_NIST = row.get("C1_NIST (J/kg-mole)", 0.0)

    # Temperature limits
    tmin_nist   = row.get("ICPtmin (K)")
    tmax_nist   = row.get("ICPtmax (K)")
    tmin_simsci = row.get("IEtmin_simsci")
    tmax_simsci = row.get("IEtmax_simsci")

    if any(pd.isna(v) for v in [
        tmin_nist, tmax_nist, tmin_simsci, tmax_simsci
    ]):
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICigas (J/kg-mole)": pd.NA,
            "ICigas_Method": "Recon",
            "ICigas_Status": "LIMITS_MISSING"
        })

    # ---------------------------------
    # Δ logic
    # ---------------------------------
    tmin = max(tmin_nist, tmin_simsci)
    tmax = min(tmax_nist, tmax_simsci)
    Delta = tmax - tmin

    # Case 1: Δ >= 0
    if Delta >= 0:
        H_nist = HIG_Trecon_NIST
        H_sim  = HIG_Trecon_SIMSCI
        recon_case = "Delta>=0"

    # Case 2: Small negative Δ
    elif abs(Delta) <= NEG_TOL:

        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HIG_Tmax_NIST
            H_sim  = HIG_Tmin_SIMSCI
            recon_case = "SmallNeg:NISTmax_SIMSCImin"

        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HIG_Tmin_NIST
            H_sim  = HIG_Tmax_SIMSCI
            recon_case = "SmallNeg:SIMSCImax_NISTmin"
        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICigas (J/kg-mole)": pd.NA,
                "ICigas_Method": "Recon",
                "ICigas_Status": "TEMP_LOGIC_ERROR"
            })

    # Case 3: Large negative Δ
    else:
        if (tmax_nist - tmin_simsci) < 0:
            H_nist = HIG_Tmax_NIST
            H_sim  = HIG_Tmax_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmax"
        elif (tmax_simsci - tmin_nist) < 0:
            H_nist = HIG_Tmin_NIST
            H_sim  = HIG_Tmin_SIMSCI
            recon_case = "LargeNeg:Use_NIST_Tmin"
        else:
            return pd.Series({
                "Delta_C1 (J/kg-mole)": pd.NA,
                "C1_Adjusted (J/kg-mole)": pd.NA,
                "ICigas (J/kg-mole)": pd.NA,
                "ICigas_Method": "Recon",
                "ICigas_Status": "TEMP_LOGIC_ERROR"
            })

    if pd.isna(H_nist) or pd.isna(H_sim):
        return pd.Series({
            "Delta_C1 (J/kg-mole)": pd.NA,
            "C1_Adjusted (J/kg-mole)": pd.NA,
            "ICigas (J/kg-mole)": pd.NA,
            "ICigas_Method": "Recon",
            "ICigas_Status": "ENTHALPY_MISSING"
        })

    # ---------------------------------
    # Final C1 adjustment
    # ---------------------------------
    R = 8.314462618
    # Delta_C1 = (H_sim  - H_nist)/R
    # C1_Adjusted = C1_NIST + Delta_C1  

    # Get equation number
    icp_eqn = row.get("ICPEqn")

    # Handle possible string/float cases
    try:
        icp_eqn = int(icp_eqn)
    except:
        icp_eqn = None

    # Apply logic
    if icp_eqn == 40:
        Delta_C1 = (H_sim - H_nist) / R
    else:
        Delta_C1 = (H_sim - H_nist)

    C1_Adjusted = C1_NIST + Delta_C1

    return pd.Series({
        "Delta_C1 (J/kg-mole)": Delta_C1,
        "C1_Adjusted (J/kg-mole)": C1_Adjusted,
        "ICigas (J/kg-mole)": C1_Adjusted,
        "ICigas_Method": f"SIMSCI-Recon[{recon_case}]",
        "ICigas_Status": "OK"
    })


# ---------------------------------------------------
# DRIVER
# ---------------------------------------------------
def calc_icp(row):
    if row["Available_in_SIMSCI"] == "No":
        return calc_icp_simsci(row)
    else:
        return calc_icp_reconciliation(row)


def main():

    excel = pd.ExcelFile(INPUT_EXCEL)

    icp_df = pd.read_excel(excel, sheet_name="ICP")
    lcp_df = pd.read_excel(excel, sheet_name="LCP")
    scp_df = pd.read_excel(excel, sheet_name="SCP")

    results = icp_df.apply(calc_icp, axis=1)
    icp_out = pd.concat([icp_df, results], axis=1)

    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        lcp_df.to_excel(writer, sheet_name="LCP", index=False)
        icp_out.to_excel(writer, sheet_name="ICP", index=False)
        scp_df.to_excel(writer, sheet_name="SCP", index=False)

    auto_adjust_column_widths(OUTPUT_EXCEL)
    print("ICP IC calculation completed successfully")


if __name__ == "__main__":
    main()

print(f"\nRuntime: {(time.time()-start):.2f} sec")
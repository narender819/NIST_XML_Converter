"""
Script:
    24_extract_trecon_propeval_results_to_excel.py

Purpose:
    This script extracts phasewise Trecon PropEval results
    from PE1 files and updates the workflow Excel sheets.

Functionality:
    - Reads cleaned phasewise Trecon workflow sheets
    - Parses NIST and SIMSCI PE1 output files
    - Extracts phasewise enthalpy values at:
        * Trecon
        * Tmin
        * Tmax
    - Updates LCP, ICP, and SCP workflow sheets
    - Tracks PE1 file availability status
    - Generates updated phasewise Excel reports
    - Auto-adjusts Excel column widths

Input:
    - Cleaned phasewise Trecon workflow Excel file
    - NIST and SIMSCI PE1 output files

Output:
    Updated Excel report containing extracted Trecon PropEval results
"""


import pandas as pd
import re
from pathlib import Path
from openpyxl import load_workbook

from utils import auto_adjust_column_widths

from config import (
    RUN_YEAR,
    OUTPUT_DIR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# OUTPUT DIRECTORIES (runtime references)
# ==================================================

PROPEVAL_RUNS_DIR = OUTPUT_DIR / "propeval_runs"

NIST_TRECON_DIR = PROPEVAL_RUNS_DIR / "NIST_Trecon"
SIMSCI_TRECON_DIR = PROPEVAL_RUNS_DIR / "SIMSCI_Trecon"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

EXCEL_PATH = PROCESSED_DIR / f"17_NIST_With_SIMSCI_Trecon_Clean.xlsx"

OUTPUT_EXCEL = PROCESSED_DIR / f"18_NIST_With_SIMSCI_Trecon_PE1Trecontoexcel.xlsx"

# ==================================================
# VALIDATION 
# ==================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(f"Missing input: {EXCEL_PATH}")

if not NIST_TRECON_DIR.exists():
    raise FileNotFoundError(f"Missing folder: {NIST_TRECON_DIR}")

if not SIMSCI_TRECON_DIR.exists():
    raise FileNotFoundError(f"Missing folder: {SIMSCI_TRECON_DIR}")


# ===================================================
# PHASE CONFIG
# ===================================================
PHASE_CONFIG = {
    "LCP": {"tag": "HL"},
    "ICP": {"tag": "HIG"},
    "SCP": {"tag": "HS"},
}

# ===================================================
# HELPERS
# ===================================================
def extract_last_float(text):
    nums = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text)
    return float(nums[-1]) if nums else None


def parse_phase_trecon(pe1_file, tag):
    """
    Extract H@Trecon, H@Tmin, H@Tmax for given phase tag (HL / HV / HS)
    """
    H_tr = H_tmin = H_tmax = None

    lines = Path(pe1_file).read_text(
        encoding="cp1252", errors="ignore"
    ).splitlines()

    for line in lines:
        u = line.upper()

        if f"{tag}@TRECON" in u:
            H_tr = extract_last_float(line)

        elif f"{tag}@TMIN" in u:
            H_tmin = extract_last_float(line)

        elif f"{tag}@TMAX" in u:
            H_tmax = extract_last_float(line)

    return H_tr, H_tmin, H_tmax


# ===================================================
# MAIN
# ===================================================
def main():

    xl = pd.ExcelFile(EXCEL_PATH)

    # Load all sheets
    sheets = {
        name: xl.parse(name)
        for name in xl.sheet_names
        if name in PHASE_CONFIG
    }

    print("Loaded sheets:", list(sheets.keys()))

    # Loop over components using TRCID as anchor
    base_df = sheets["LCP"][["TRCID", "Aliases", "LibraryID"]].drop_duplicates()

    for _, base in base_df.iterrows():

        trcid = int(base["TRCID"])
        alias = str(base["Aliases"]).strip()
        libid = str(base["LibraryID"]).strip() if pd.notna(base["LibraryID"]) else None

        pe1_nist = Path(NIST_TRECON_DIR) / f"{alias}_{trcid}_NIST.pe1"
        pe1_sim  = (
            Path(SIMSCI_TRECON_DIR) / f"{libid}_{libid}_SIMSCI.pe1"
            if libid else None
        )

        # ------------------------------------------------
        # ROUTE DATA PHASE-WISE
        # ------------------------------------------------
        for phase, cfg in PHASE_CONFIG.items():

            dfp = sheets[phase]
            mask = dfp["TRCID"] == trcid
            if not mask.any():
                continue

            tag = cfg["tag"]

            # ---------- NIST ----------
            if pe1_nist.exists():
                Htr, Hmin, Hmax = parse_phase_trecon(pe1_nist, tag)

                dfp.loc[mask, f"{tag}@Trecon_NIST (J/kg-mole)"] = Htr
                dfp.loc[mask, f"{tag}@Tmin_NIST (J/kg-mole)"] = Hmin
                dfp.loc[mask, f"{tag}@Tmax_NIST (J/kg-mole)"] = Hmax
                dfp.loc[mask, "NIST_Trecon_PE1_Found"] = "Yes"
            else:
                dfp.loc[mask, "NIST_Trecon_PE1_Found"] = "No"

            # ---------- SIMSCI ----------
            if pe1_sim and pe1_sim.exists():
                Htr, Hmin, Hmax = parse_phase_trecon(pe1_sim, tag)

                dfp.loc[mask, f"{tag}@Trecon_SIMSCI (J/kg-mole)"] = Htr
                dfp.loc[mask, f"{tag}@Tmin_SIMSCI (J/kg-mole)"] = Hmin
                dfp.loc[mask, f"{tag}@Tmax_SIMSCI (J/kg-mole)"] = Hmax
                dfp.loc[mask, "SIMSCI_Trecon_PE1_Found"] = "Yes"
            else:
                dfp.loc[mask, "SIMSCI_Trecon_PE1_Found"] = "No"

        print(f" Parsed PE1 TRCID {trcid}")

    # ------------------------------------------------
    # WRITE BACK TO EXCEL
    # ------------------------------------------------
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="w") as writer:
        for name, df_out in sheets.items():
            df_out.to_excel(writer, sheet_name=name, index=False)

    auto_adjust_column_widths(OUTPUT_EXCEL)
    print(f"\n Results written to: {OUTPUT_EXCEL}")



# ===================================================
if __name__ == "__main__":
    main()

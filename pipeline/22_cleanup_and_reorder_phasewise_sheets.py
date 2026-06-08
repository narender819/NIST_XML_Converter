"""
Script:
    22_cleanup_and_reorder_phasewise_sheets.py

Purpose:
    This script cleans, organizes, and reorders phasewise
    thermodynamic workflow sheets for LCP, ICP, and SCP data.

Functionality:
    - Reads phasewise workflow Excel sheets
    - Reorders columns using predefined phasewise layouts
    - Organizes identity, constants, ranges, flags,
      and Trecon calculation columns
    - Removes unnecessary columns from output sheets
    - Generates cleaned and standardized workflow sheets
    - Auto-adjusts Excel column widths

Input:
    Phasewise Trecon workflow Excel file

Output:
    Cleaned and reordered phasewise workflow Excel report
"""


import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

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
INPUT_EXCEL = (
    PROCESSED_DIR
    / f"16_NIST_With_SIMSCI_Trecon.xlsx"
)

OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"17_NIST_With_SIMSCI_Trecon_Clean.xlsx"
)


# ---------------------------------------------------
# COLUMN ORDER DEFINITIONS
# ---------------------------------------------------

IDENTITY_COLS = [
    "TRCID",
    "CASNO",
    "ComponentName",
    "Aliases",
    "LibraryID",
]

FLAG_COLS = [
    "Available_in_SIMSCI",
    "LCP_Exists",
    "ICP_Exists",
    "SCP_Exists",
    "HVAP_Exists"
]

PHASE_ORDER = {
    "LCP": {
        "ranges": [
            "LCPtmin (K)",
            "LCPtmax (K)",
            "LEtmin_simsci",
            "LEtmax_simsci",
            "HL@0C (J/kg-mole)",
            "HL@Tmin (J/kg-mole)",
            "HL@Tmax (J/kg-mole)",
            "HL@Tnbp (J/kg-mole)",
            "HL@Tnmp (J/kg-mole)",
            "HL@NBP (J/kg-mole)",
            "HL@NMP (J/kg-mole)",     
            "dHL/dT@Tmin",
            "dHL/dT@Tmax"          

        ],
        "trecon": ["LCP_Trecon_SIMSCI (K)", "LCP_Trecon_NIST (K)", "LCP_Trecon_Remark"],
        "constants": [
            "NBP (K)",
            "NMP (K)",
            "TC (K)",
            "HFUSIONNMP (kJ/kg-mol)",
        ],
    },
    "ICP": {
        "ranges": [
            "ICPtmin (K)",
            "ICPtmax (K)",
            "IEtmin_simsci",
            "IEtmax_simsci",
            "ICPEqn",
            "HIG@NBP (J/kg-mole)",
            "HV@NBP(J/kg-mole)",
            "HIG@Tnbp (J/kg-mole)",
            "HIG@Tmax (J/kg-mole)",
            "HVAP@NBP(J/kg-mole)",
            "HVAP@Tnbp(J/kg-mole)",
            "HVAP@Mid (J/kg-mole)",
            "HDeparture (J/kg-mole)",
            "HL@Tnbp (J/kg-mole)",
            "HVAPtmin (K)",
            "HAVAPtmax (K)",
            "LCPtmin (K)",
            "LCPtmax (K)",
            "HL@Tmin (J/kg-mole)",
            "HL@Tmax (J/kg-mole)",
            "HL@Tnbp (J/kg-mole)",
            "dHL/dT@Tmin",
            "dHL/dT@Tmax"
            
        ],
        "trecon": ["ICP_Trecon_SIMSCI (K)", "ICP_Trecon_NIST (K)", "ICP_Trecon_Remark"],
        "constants": [
            "NBP (K)",
            "TC (K)",
        ],
    },
    "SCP": {
        "ranges": [
            "SCPtmin (K)",
            "SCPtmax (K)",
            "SEtmin_simsci",
            "SEtmax_simsci",
            "SCP_Tnmp (K)",
            "HS@Tnmp (J/kg-mole)",
            "HS@Tmax (J/kg-mole)",
            "HL@Tnmp (J/kg-mole)",
            "dHL/dT@Tmin",
            "dHL/dT@Tmax",
            "HL@Tmin (J/kg-mole)",
            "HL@Tmax (J/kg-mole)",
            "HL@Tnbp (J/kg-mole)"
 
        ],
        "trecon": ["SCP_Trecon_SIMSCI (K)", "SCP_Trecon_NIST (K)", "SCP_Trecon_Remark"],
        "constants": [
            "NBP (K)",
            "NMP (K)",
            "TC (K)",
            "HFUSIONNMP (kJ/kg-mol)",
        ],
    },
}

# ---------------------------------------------------
# HELPERS
# ---------------------------------------------------
def existing(df, cols):
    """Return only columns that actually exist in df"""
    return [c for c in cols if c in df.columns]


def clean_and_order_sheet(df, phase):
    cfg = PHASE_ORDER[phase]

    ordered_cols = (
        existing(df, IDENTITY_COLS)
        + existing(df, cfg["constants"])        
        + existing(df, cfg["ranges"])
        + existing(df, FLAG_COLS)
        + existing(df, cfg["trecon"])
        
    )

    return df[ordered_cols].copy()


# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
def main():
    xls = pd.ExcelFile(INPUT_EXCEL)
    writer = pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl")

    for sheet_name in ["LCP", "ICP", "SCP"]:
        if sheet_name not in xls.sheet_names:
            print(f"Sheet not found, skipped: {sheet_name}")
            continue

        df = pd.read_excel(xls, sheet_name)
        clean_df = clean_and_order_sheet(df, sheet_name)

        clean_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(
            f"{sheet_name}: reordered to {len(clean_df.columns)} columns"
        )

    writer.close()
    # After writer.close()
    auto_adjust_column_widths(OUTPUT_EXCEL)
    print("\n Housekeeping + column ordering complete")
    print(f" Output written to:\n{OUTPUT_EXCEL}")


# ---------------------------------------------------
if __name__ == "__main__":
    main()

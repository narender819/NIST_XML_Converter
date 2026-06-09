"""
Script:
    34_update_xml_enthalpy_integrated_constants.py

Purpose:
    This script updates XML enthalpy correlation C1 values
    using calculated integrated constants for LCP, ICP, and SCP.

Functionality:
    - Reads calculated IC workflow Excel sheets
    - Processes generated NIST XML component files
    - Updates LiquidEnthalpy C1 parameters
    - Updates IdealEnthalpy C1 parameters
    - Updates SolidEnthalpy C1 parameters
    - Uses adjusted C1 values when available
    - Falls back to IC values when adjusted C1 is unavailable
    - Generates updated XML component files

Input:
    - Integrated constant calculation Excel file
    - Generated NIST XML component files

Output:
    Updated XML component files containing revised enthalpy C1 values
"""



import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    XML_DIR,
    XML_LIBRARY_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# INPUT FILE
# ==================================================

INPUT_EXCEL = PROCESSED_DIR / f"23_NIST_ICCalc_SCP_LCP_ICP.xlsx"

# ==================================================
# XML DIRECTORIES
# ==================================================

XML_INPUT_DIR = XML_LIBRARY_DIR / "02_missingprop_refcode9999"

OUTPUT_XML_DIR = XML_LIBRARY_DIR / "03_iccalc_updated_C1"

# Create only new output folder
OUTPUT_XML_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# VALIDATION 
# ==================================================

if not INPUT_EXCEL.exists():
    raise FileNotFoundError(f"Missing input: {INPUT_EXCEL}")

if not XML_INPUT_DIR.exists():
    raise FileNotFoundError(f"Missing XML folder: {XML_INPUT_DIR}")

# ==================================================
# DEBUG (optional)
# ==================================================

print("Input Excel:", INPUT_EXCEL)
print("XML input dir:", XML_INPUT_DIR)
print("XML output dir:", OUTPUT_XML_DIR)


# ---------------------------------------------------
# LOAD EXCEL SHEETS
# ---------------------------------------------------
df_lcp = pd.read_excel(INPUT_EXCEL, sheet_name="LCP")
df_icp = pd.read_excel(INPUT_EXCEL, sheet_name="ICP")
df_scp = pd.read_excel(INPUT_EXCEL, sheet_name="SCP")

# Convert sheets to lookup dictionaries
lcp_dict = {r["Aliases"]: r for _, r in df_lcp.iterrows()}
icp_dict = {r["Aliases"]: r for _, r in df_icp.iterrows()}
scp_dict = {r["Aliases"]: r for _, r in df_scp.iterrows()}


# ---------------------------------------------------
# UPDATE FUNCTION
# ---------------------------------------------------
def update_c1(root, corr_name, new_c1):
    for corr in root.findall(".//correlation"):
        if corr.get("name") == corr_name:
 
            params = corr.find("parameters")

            if params is not None and params.text:
                values = params.text.split()
                values[0] = f"{new_c1:.10f}"
                params.text = " ".join(values)
                return True

            elif corr.text and corr.text.strip():
                values = corr.text.split()
                values[0] = f"{new_c1:.10f}"
                corr.text = " ".join(values)
                return True

    return False


# ---------------------------------------------------
# PROCESS XML FILES
# ---------------------------------------------------
processed = 0

for xml_path in XML_INPUT_DIR.glob("*.xml"):

    alias = xml_path.stem.replace("Comp-NIST-", "")

    tree = ET.parse(xml_path)
    root = tree.getroot()

    updated_any = False

    # ---------------- LCP ----------------
    row = lcp_dict.get(alias)
    if row is not None and row.get("LCP_Exists") == "Yes":
        c1 = row["C1_Adjusted (J/kg-mole)"] if pd.notna(row["C1_Adjusted (J/kg-mole)"]) else row["ICliq (J/kg-mole)"]
        if pd.notna(c1):
            if update_c1(root, "LiquidEnthalpy", float(c1)):
                print(f"LCP updated | {alias}")
                updated_any = True

    # ---------------- ICP ----------------
    row = icp_dict.get(alias)
    if row is not None and row.get("ICP_Exists") == "Yes":
        c1 = row["C1_Adjusted (J/kg-mole)"] if pd.notna(row["C1_Adjusted (J/kg-mole)"]) else row["ICigas (J/kg-mole)"]
        if pd.notna(c1):
            if update_c1(root, "IdealEnthalpy", float(c1)):
                print(f"ICP updated | {alias}")
                updated_any = True

    # ---------------- SCP ----------------
    row = scp_dict.get(alias)
    if row is not None and row.get("SCP_Exists") == "Yes":
        c1 = row["C1_Adjusted (J/kg-mole)"] if pd.notna(row["C1_Adjusted (J/kg-mole)"]) else row["ICSolid (J/kg-mole)"]
        if pd.notna(c1):
            if update_c1(root, "SolidEnthalpy", float(c1)):
                print(f"SCP updated | {alias}")
                updated_any = True

    # ---------------- WRITE XML ----------------
    out_path = OUTPUT_XML_DIR / xml_path.name
    tree.write(out_path, encoding="utf-8", xml_declaration=True)

    processed += 1


# ---------------------------------------------------
# SUMMARY
# ---------------------------------------------------
print("\n===============================")
print(f"Total XML processed : {processed}")
print(f"Input XML count     : {len(list(XML_INPUT_DIR.glob('*.xml')))}")
print(f"Output XML count    : {len(list(OUTPUT_XML_DIR.glob('*.xml')))}")
print("===============================")
"""
Script:
    33_update_xml_missing_property_refcodes.py

Purpose:
    This script updates XML property refCodes for selected
    missing thermodynamic properties using predefined rules.

Functionality:
    - Reads missing property status Excel data
    - Processes generated NIST XML component files
    - Identifies properties marked as missing
    - Updates XML property refCode values to 9999
    - Applies updates for selected properties such as:
        * TC
        * PC
        * ACENTRIC
    - Generates updated XML component files

Input:
    - Missing property status Excel file
    - Generated NIST XML component files

Output:
    Updated XML component files with modified property refCodes
"""


import pandas as pd
import xml.etree.ElementTree as ET
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

XML_BASE_DIR = (
    OUTPUT_DIR
    / "xml"
)

# ==================================================
# INPUT FILES
# ==================================================
INPUT_EXCEL = (
    PROCESSED_DIR
    / f"11_Missing_Fillin_Properties_{RUN_YEAR}.xlsx"
)

# ==================================================
# XML INPUT / OUTPUT
# ==================================================
XML_DIR = (
    XML_BASE_DIR
    / "Libraryfiles_NIST"
    / "01_generated"
)

OUTPUT_XML_DIR = (
    XML_BASE_DIR
    / "Libraryfiles_NIST"
    / "02_missingprop_refcode9999"
)

OUTPUT_XML_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ---------------- LOAD EXCEL ----------------
df = pd.read_excel(INPUT_EXCEL)

# Convert to dictionary using trcid
excel_dict = {str(r["trcid"]): r for _, r in df.iterrows()}

# PROPERTIES = [
#     "NMP", "TC", "PC", "SG60F", "MVOL25C",
#     "VC", "PTP", "TTP", "HVAPNBP", "ACENTRIC"
# ]

PROPERTIES = [
    "TC", "PC","ACENTRIC"
]
# ---------------- PROCESS XML ----------------
processed = 0

for xml_path in XML_DIR.glob("*.xml"):

    alias = xml_path.stem.replace("Comp-NIST-", "")  # NST5523
    trcid = alias.replace("NST", "")                 # 5523

    row = excel_dict.get(trcid)

    if row is None:
        continue

    tree = ET.parse(xml_path)
    root = tree.getroot()

    updated_any = False

    # -------- PROPERTY TAGGING --------
    for prop in PROPERTIES:

        if row[prop] == "MISSING":

            for p in root.findall(".//property"):

                if p.get("name") == prop:

                    value = p.get("value")

                    # Only tag if value exists
                    if value is not None and value.strip() != "":
                        p.set("refCode", "9999")
                        print(f"{prop} tagged | {alias}")
                        updated_any = True

    # -------- WRITE OUTPUT --------
    out_path = OUTPUT_XML_DIR / xml_path.name
    tree.write(out_path, encoding="utf-8", xml_declaration=True)

    processed += 1

# ---------------- SUMMARY ----------------
print("\n===============================")
print(f"Total XML processed : {processed}")
print("===============================")
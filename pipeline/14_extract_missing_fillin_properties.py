"""
Script:
    14_extract_missing_fillin_properties.py

Purpose:
    This script identifies missing fill-in thermodynamic properties
    from processed component JSON files and generates a consolidated
    Excel report.

Functionality:
    - Reads processed component JSON files
    - Checks required thermodynamic properties
    - Identifies missing property values
    - Extracts component details such as:
        * TRCID
        * Component Name
        * CAS Number
        * Source Folder
    - Generates a wide-format Excel report showing missing properties

Input:
    Processed component JSON files

Output:
    Excel report containing missing fill-in property information
"""

import os
import json
import pandas as pd

from pathlib import Path

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# INPUT FOLDERS
# ==================================================

COMPONENTS_MASTER_DIR = PROCESSED_DIR / "1_components_Inmaster_withsimsciid"
COMPONENTS_ASSIGNED_DIR = PROCESSED_DIR / "3_components_notInmaster_assignedsimsciid"

folders = [
    COMPONENTS_MASTER_DIR,
    COMPONENTS_ASSIGNED_DIR
]


PROPERTIES = [
    "NMP", "TC", "PC", "SG60F", "MVOL25C",
    "VC", "PTP", "TTP", "HVAPNBP", "ACENTRIC"
]

output_data = []

for folder in folders:
    for file in os.listdir(folder):
        if file.endswith(".json"):

            file_path = os.path.join(folder, file)

            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                props = data.get("properties", {})

                #  FIXED TRCID EXTRACTION
                trcid = file.split("_")[0]

                name = data.get("name")
                cas = data.get("CASNO")

                row = {
                    "trcid": trcid,
                    "Name": name,
                    "CAS": cas,
                    "SourceFolder": os.path.basename(folder)
                }

                for prop in PROPERTIES:
                    row[prop] = "MISSING" if props.get(prop) is None else ""

                output_data.append(row)

            except Exception as e:
                print(f"Error reading {file}: {e}")

df = pd.DataFrame(output_data)

cols = ["trcid", "Name", "CAS"] + PROPERTIES + ["SourceFolder"]
df = df[cols]

output_path = (PROCESSED_DIR / f"11_Missing_Fillin_Properties_{RUN_YEAR}.xlsx")
df.to_excel(output_path, index=False)

print("Excel generated at:", output_path)
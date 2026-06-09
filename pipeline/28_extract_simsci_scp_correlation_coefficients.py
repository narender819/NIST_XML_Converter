"""
Script:
    28_extract_simsci_scp_correlation_coefficients.py

Purpose:
    This script extracts Solid Enthalpy (SCP) correlation
    coefficients and metadata from existing SIMSCI XML libraries.

Functionality:
    - Reads SIMSCI XML library component files
    - Extracts component identification details
    - Extracts SolidEnthalpy correlation information
    - Extracts SCP equation numbers and temperature ranges
    - Extracts SCP correlation coefficients
    - Applies library priority-based component selection
    - Removes duplicate CAS-based entries
    - Generates a consolidated SIMSCI SCP coefficient Excel report

Input:
    SIMSCI XML library component files

Output:
    Excel report containing SIMSCI SCP correlation coefficients and metadata
"""


import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path
import re
from openpyxl import load_workbook


from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    LIBRARIES_INPUT_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITES (derived)
# ==================================================

SIMSCI_XML_ROOT = LIBRARIES_INPUT_DIR / "SIMSCI" / "Libraries_xmlfiles_2025"

# ==================================================
# OUTPUT FILE
# ==================================================

OUTPUT_EXCEL = PROCESSED_DIR / f"19_NIST_SIMSCI_SCP_Key.xlsx"

# =========================================================
# LIBRARY PRIORITY
# =========================================================
LIB_PRIORITY = {
    "Edlib": 1,
    "DIPPR": 2,
    "BioLib": 3,
    "Chemical": 4,
    "Mining": 5,
}


# =========================================================
# HELPERS
# =========================================================

def normalize_cas(cas_str):
    """
    Normalize CAS to pure numeric format.
    Example:
        71-43-2      -> 71432
        71432        -> 71432
        000071-43-2  -> 71432
    """
    if not cas_str:
        return None

    # Remove dashes and leading zeros
    cas_clean = cas_str.strip().replace("-", "").lstrip("0")

    return cas_clean



def extract_numeric_coefficients(text):
    if not text:
        return []

    return re.findall(r'[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?', text)


# =========================================================
# EXTRACTION FUNCTION
# =========================================================

def extract_solid_enthalpy(xml_path, priority):

    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()

        id_node = root.find("id")
        if id_node is None:
            return None

        cas = normalize_cas(id_node.get("casNum"))
        comp_name = id_node.get("name")
        aliases = id_node.text.strip() if id_node.text else ""

        corr = root.find(".//correlation[@name='SolidEnthalpy']")
        if corr is None:
            return None

        raw_text = corr.text if corr.text else ""
        coeffs = extract_numeric_coefficients(raw_text)

        data = {
            "CASNO": cas,
            "ComponentName": comp_name,
            "Aliases": aliases,
            "SCPEqn_sim": int(corr.get("equation")) if corr.get("equation") else None,
            "SEtmin_simsci": float(corr.get("tMin")) if corr.get("tMin") else None,
            "SEtmax_simsci": float(corr.get("tMax")) if corr.get("tMax") else None,
            "PropUnit_sim": corr.get("propUnit"),
            "Priority": priority
        }

        # Populate C1–C8
        for i in range(1, 9):
            if i <= len(coeffs):
                try:
                    data[f"C{i}_sim"] = float(coeffs[i-1])
                except:
                    data[f"C{i}_sim"] = None
            else:
                data[f"C{i}_sim"] = None

        return data

    except Exception as e:
        print(f"Error in {xml_path.name}: {e}")
        return None


# =========================================================
# MAIN ENGINE
# =========================================================

records = []

for xml_file in Path(SIMSCI_XML_ROOT).rglob("*.xml"):

    filename = xml_file.name.lower()

    # Skip Bank files globally
    if filename.startswith("bank") or filename.startswith("__bank"):
        continue

    path_parts = [p.lower() for p in xml_file.parts]

    library_found = None
    for lib in LIB_PRIORITY.keys():
        if lib.lower() in path_parts:
            library_found = lib
            break

    if not library_found:
        continue

    # Special restriction for Edlib
    if library_found.lower() == "edlib":
        try:
            ed_index = path_parts.index("edlib")
            subfolder = path_parts[ed_index + 1]
        except:
            continue

        if subfolder not in ["simsci", "process"]:
            continue

    priority = LIB_PRIORITY[library_found]

    result = extract_solid_enthalpy(xml_file, priority)

    if result:
        records.append(result)


if not records:
    print("No valid records found.")
else:

    df = pd.DataFrame(records)

    # Apply priority ranking
    df.sort_values(by=["CASNO", "Priority"], inplace=True)
    df_clean = df.drop_duplicates(subset=["CASNO"], keep="first")
    df_clean = df_clean.drop(columns=["Priority"])

    final_cols = [
        "CASNO", "ComponentName", "Aliases",
        "SCPEqn_sim",
        "SEtmin_simsci", 
        "SEtmax_simsci",
        "PropUnit_sim",
        "C1_sim", "C2_sim", "C3_sim", "C4_sim",
        "C5_sim", "C6_sim", "C7_sim", "C8_sim"
    ]

    df_clean = df_clean.reindex(columns=final_cols)

    # Append mode writing
    try:
        with pd.ExcelWriter(
            OUTPUT_EXCEL,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            df_clean.to_excel(writer, sheet_name="SIMSCI_SCP_Key", index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="w") as writer:
            df_clean.to_excel(writer, sheet_name="SIMSCI_SCP_Key", index=False)

    print(f"Completed. Extracted {len(df_clean)} unique components.")
    print("Sheet written: SIMSCI_SCP_Key")

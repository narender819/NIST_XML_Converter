"""
Script:
    05_extract_dippr_family_data.py

Purpose:
    This script reads NIST component JSON files, extracts component,
    alias, and DIPPR family information, and generates a consolidated
    Excel report.

Functionality:
    - Reads component JSON files from the configured folder
    - Extracts:
        * TRCID
        * CASRN
        * Component Name
        * All Aliases
        * Formula
        * DIPPR Family
        * DIPPR Subfamily
    - Extracts CASRN from JSON file names
    - Consolidates all extracted records into a single Excel report

Input:
    Component JSON files generated from the extraction process

Output:
    Excel report containing DIPPR family and component information
"""

from pathlib import Path
import os
import json
import pandas as pd
import re

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR

PROCESSED_DIR = OUTPUT_DIR / "processed" / "full_library"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================
JSON_FOLDER = OUTPUT_DIR / "json"

OUTPUT_FILE = (
    PROCESSED_DIR
    / "4_DIPPR_Family_Extraction_All_json.xlsx"
)


def extract_cas_from_filename(filename):
    """
    Extract CAS number from filename format: *_<CAS>.json
    Example: 16_CASNO_2207014.json → 2207014
    """
    match = re.search(r'_(\d+)\.json$', filename)
    return match.group(1) if match else "NA"


def process_folder(folder_path):
    """
    Reads every JSON file in the folder and extracts:
    TRCID, Component Name, Aliases, Formula, DIPPR Family, Subfamily, CAS number
    """
    records = []

    for file in os.listdir(folder_path):
        if not file.lower().endswith(".json"):
            continue

        file_path = os.path.join(folder_path, file)

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            print(f"Error reading {file}: {e}")
            continue

        compounds = data.get("compounds", [])
        if not compounds:
            print(f"No compounds found in {file}")
            continue

        for comp in compounds:

            trcid = comp.get("trcid", "NA")

            formula = comp.get("formula", "NA")

            names_list = comp.get("names", [])
            component_name = names_list[0] if names_list else "NA"
            all_aliases = ", ".join(names_list) if names_list else "NA"

            dippr = comp.get("DIPPR", {})
            dippr_family = dippr.get("family", "NA")
            dippr_subfamily = dippr.get("subfamily", "NA")

            cas_number = extract_cas_from_filename(file)

            records.append({
                "TRCID": trcid,
                "CASRN": cas_number,
                "Component_Name": component_name,
                "All_Aliases": all_aliases,
                "Formula": formula,
                "DIPPR_Family": dippr_family,
                "DIPPR_Subfamily": dippr_subfamily
            })

    return records


def main():

    # Extract records
    all_records = process_folder(JSON_FOLDER)

    # Convert to DataFrame
    df = pd.DataFrame(all_records)

    # Write Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="all_data", index=False)

    print(f"Excel file created: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
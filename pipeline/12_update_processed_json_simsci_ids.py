"""
Script:
    11a_assign_simsci_ids_to_unmatched_json.py

Purpose:
    This script updates processed component JSON files with
    assigned SIMSCI IDs using TRCID and CASRN mapping data.

Functionality:
    - Reads the SIMSCI assignment Excel dataset
    - Builds TRCID + CASRN to SIMSCI ID mapping
    - Processes component JSON files
    - Updates missing SIMSCIID values in JSON files
    - Skips files with existing SIMSCIID values
    - Generates updated JSON files in the output folder
    - Displays processing summary and statistics

Input:
    - SIMSCI assignment Excel file
    - Processed component JSON files

Output:
    Updated JSON files containing assigned SIMSCIID values
"""

import os
import json
import pandas as pd

from pathlib import Path

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
excel_file = (
    PROCESSED_DIR
    / "9_DIPPR_Family_Extraction_All_With_Smiles_none_Predicted_Family_SIMSCI_ASSIGNED.xlsx"
)

json_input_folder = (
    PROCESSED_DIR
    / "2_components_notInmaster_nosimsciid"
)

json_output_folder = (
    PROCESSED_DIR
    / "3_components_notInmaster_assignedsimsciid"
)

json_output_folder.mkdir(
    parents=True,
    exist_ok=True
)


# ==============================
# LOAD EXCEL MAPPING (TRCID + CAS)
# ==============================
df = pd.read_excel(excel_file)

df["CASRN"] = df["CASRN"].astype(str).str.strip()
df["TRCID"] = df["TRCID"].astype(str).str.strip()
df["SIMSCI_ID"] = pd.to_numeric(df["SIMSCI_ID"], errors="coerce")

# --- Optional Safety Check ---
duplicates = df[df.duplicated(subset=["TRCID", "CASRN"], keep=False)]
if not duplicates.empty:
    print("WARNING: Duplicate TRCID + CAS combinations found in Excel!")
    print(duplicates)

# --- Create Composite Mapping ---
mapping = {}

for _, row in df.iterrows():
    key = f"{row['TRCID']}_{row['CASRN']}"
    mapping[key] = row["SIMSCI_ID"]

# ==============================
# STATS
# ==============================
total = 0
updated = 0
skipped_no_cas = 0
skipped_no_id = 0
skipped_existing = 0
errors = 0

# ==============================
# PROCESS JSON FILES
# ==============================
for fname in os.listdir(json_input_folder):

    if not fname.lower().endswith(".json"):
        continue

    total += 1
    json_path = os.path.join(json_input_folder, fname)

    try:
        # --- Validate filename format ---
        # Expected format: TRCID_CASNO_<CAS>.json
        if "_CASNO_" not in fname:
            skipped_no_cas += 1
            continue

        # --- Extract TRCID ---
        trcid = fname.split("_")[0].strip()

        # --- Extract CAS ---
        cas = fname.split("_CASNO_")[1].replace(".json", "").strip()

        key = f"{trcid}_{cas}"
        simsci_id = mapping.get(key)

        if simsci_id is None or pd.isna(simsci_id):
            skipped_no_id += 1
            continue

        # --- Load JSON ---
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # --- Validate SIMSCIID field ---
        if "SIMSCIID" not in data:
            skipped_existing += 1
            continue

        if data["SIMSCIID"] is not None:
            skipped_existing += 1
            continue

        # --- Update SIMSCIID ---
        data["SIMSCIID"] = int(simsci_id)

        # --- Write to output folder ---
        out_path = os.path.join(json_output_folder, fname)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        updated += 1

    except Exception as e:
        print(f"ERROR processing {fname}: {e}")
        errors += 1

# ==============================
# SUMMARY
# ==============================
print("\n===== SIMSCI JSON UPDATE SUMMARY =====")
print(f"Total JSON files scanned      : {total}")
print(f"Successfully updated          : {updated}")
print(f"Skipped (CAS/TRCID issue)     : {skipped_no_cas}")
print(f"Skipped (No match in Excel)   : {skipped_no_id}")
print(f"Skipped (SIMSCIID exists)     : {skipped_existing}")
print(f"Errors                        : {errors}")
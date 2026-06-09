"""
Script:
    25_append_master_alias_names_to_json.py

Purpose:
    This script appends standardized alias names from the
    master component list into processed component JSON files.

Functionality:
    - Reads the master component Excel file
    - Extracts primary alias names using CAS matching
    - Processes component JSON files
    - Generates standardized aliases using:
        * TRCID-based alias
        * Master alias name
    - Updates JSON alias fields
    - Generates updated JSON files
    - Creates alias assignment status reports

Input:
    - Master component Excel file
    - Processed component JSON files

Output:
    - Updated JSON files containing standardized aliases
    - Alias assignment Excel report
"""



import os
import json
import pandas as pd
from pathlib import Path

from pathlib import Path
from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITES
# ==================================================

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

MASTER_FILE = EXCEL_INPUT_DIR / "5_Master_Component_List.xlsx"

# ==================================================
# INPUT / OUTPUT DIRECTORIES
# ==================================================

JSON_INPUT_DIR = PROCESSED_DIR / "1_components_Inmaster_withsimsciid_fillin"

OUTPUT_DIR = PROCESSED_DIR / "1_components_Inmaster_withsimsciid_fillin_alias_updated"

#  Create only new output subfolder
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# OUTPUT FILE
# ==================================================

REPORT_PATH = OUTPUT_DIR / "25_alias_assignment_report.xlsx"

# ==================================================
# VALIDATION 
# ==================================================

if not MASTER_FILE.exists():
    raise FileNotFoundError(f"Missing input: {MASTER_FILE}")

if not JSON_INPUT_DIR.exists():
    raise FileNotFoundError(f"Missing folder: {JSON_INPUT_DIR}")


# ---------------- HELPER ----------------
def normalize_cas(cas):
    if pd.isna(cas):
        return None
    digits = ''.join(ch for ch in str(cas) if ch.isdigit())
    return digits.lstrip('0') or '0'

# ---------------- STEP 1: READ MASTER ----------------
print("Loading MASTER file...")

master_df = pd.read_excel(MASTER_FILE, header=5)
master_df.columns = master_df.columns.str.strip()

# Normalize CAS
master_df["_CAS_NORM"] = master_df["CAS Number"].apply(normalize_cas)

# Extract primary alias (FIRST token)
def get_primary_alias(alias_str):
    if pd.isna(alias_str):
        return None
    parts = str(alias_str).split()
    return parts[0] if parts else None

master_df["PRIMARY_ALIAS"] = master_df["All aliases, tokenized, UNIQUE entries"].apply(get_primary_alias)

# Build lookup
cas_to_alias = dict(zip(master_df["_CAS_NORM"], master_df["PRIMARY_ALIAS"]))

print(f"Total aliases loaded: {len(cas_to_alias)}")

# ---------------- STEP 2: PROCESS JSON ----------------
alias_log = []

json_files = list(Path(JSON_INPUT_DIR).glob("*.json"))

print(f"JSON files found: {len(json_files)}")

if len(json_files) == 0:
    print("[WARNING] No JSON files found in input directory!")

print(f"Processing {len(json_files)} JSON files...")

processed_count = 0
skipped_count = 0

for json_file in json_files:

    filename = json_file.name

    # -------- SAFE FILENAME PARSING --------
    if "_CASNO_" not in filename:
        print(f"[SKIPPED] Invalid filename format: {filename}")
        skipped_count += 1
        continue

    try:
        parts = filename.split("_CASNO_")
        trcid = parts[0]
        cas_raw = parts[1].replace(".json", "")
    except Exception as e:
        print(f"[ERROR] Parsing failed for {filename}: {e}")
        skipped_count += 1
        continue

    cas_norm = normalize_cas(cas_raw)

    print(f"[DEBUG] Processing: {filename} : TRCID: {trcid}, CAS: {cas_norm}")

    # -------- READ JSON --------
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERROR] Failed to read JSON {filename}: {e}")
        skipped_count += 1
        continue

    # ---------------- ALIAS LOGIC ----------------
    base_alias = f"NST{trcid}"

    simsci_alias = cas_to_alias.get(cas_norm)

    if simsci_alias and str(simsci_alias).strip():
        final_alias = f"{base_alias} {simsci_alias}"
        status = "Matched"
    else:
        final_alias = base_alias
        status = "Missing"

    # Assign alias
    data["aliases"] = [final_alias]

    # -------- SAVE UPDATED JSON --------
    try:
        output_file = (OUTPUT_DIR/ filename)
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"[ERROR] Failed to write JSON {filename}: {e}")
        skipped_count += 1
        continue

    # -------- LOG --------
    alias_log.append({
        "TRCID": trcid,
        "CAS": cas_norm,
        "Final_Alias": final_alias,
        "Primary_Alias_From_Master": simsci_alias if simsci_alias else "",
        "Status": status
    })

    processed_count += 1

# ---------------- STEP 3: SAVE REPORT ----------------
report_df = pd.DataFrame(alias_log)

if not report_df.empty:
    report_df.to_excel(REPORT_PATH, index=False)
else:
    print("[WARNING] No data to write in report.")

# ---------------- SUMMARY ----------------
print("\n===== SUMMARY =====")

print(f"Total files found      : {len(json_files)}")
print(f"Successfully processed: {processed_count}")
print(f"Skipped               : {skipped_count}")

if not report_df.empty and "Status" in report_df.columns:
    print("\nStatus Breakdown:")
    print(report_df["Status"].value_counts())
else:
    print("\nNo records processed. alias_log is empty.")

print(f"\nReport saved at: {REPORT_PATH}")
print("Processing completed.")
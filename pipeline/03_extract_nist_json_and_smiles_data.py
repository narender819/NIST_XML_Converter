
"""
Script:
    2_extract_json_smiles.py

Purpose:
    This script extracts component JSON data and SMILES information
    from NIST SQLite databases and generates structured output files
    for further processing and analysis.

Functionality:
    - Connects to the NIST PURE and COMPOUNDS SQLite databases
    - Extracts component JSON data using TRCID and CASRN
    - Validates and exports JSON data into individual JSON files
    - Logs unmatched or invalid records into an Excel report
    - Extracts SMILES data from the COMPOUNDS database
    - Generates an Excel report containing TRCID, CASRN, and SMILES data

Input:
    - Pure.models.sqlite database
    - Compounds.sqlite database

Output:
    - Individual JSON files for each valid component
    - Excel report for unmatched/skipped components
    - Excel report containing extracted SMILES data
"""

import sqlite3
import os
import json
import pandas as pd
from openpyxl import Workbook

from pathlib import Path


from config import (
    RUN_YEAR,
    BASE_DIR,
    OUTPUT_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# DATABASES
# ==================================================

DB_DIR = BASE_DIR / "db" / RUN_YEAR

DB_PATH = DB_DIR / "Pure.models.sqlite"
DB_PATH_SMILES = DB_DIR / "Compounds.sqlite"

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================

JSON_OUTPUT_DIR = OUTPUT_DIR / "json"
SMILES_OUTPUT_DIR = OUTPUT_DIR / "smiles"

# ==================================================
# OUTPUT FILES
# ==================================================

UNMATCHED_FILE = OUTPUT_DIR / "unmatched_components.xlsx"

OUTPUT_FILE_SMILES = SMILES_OUTPUT_DIR / f"1_compounds_smiles_{RUN_YEAR}.xlsx"

# # Ensure output folder exists
os.makedirs(JSON_OUTPUT_DIR, exist_ok=True)
os.makedirs(SMILES_OUTPUT_DIR, exist_ok=True)

def extract_and_log():
    print("Connecting to the NIST database...")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # query = "SELECT TRCID, CASRN, JSON FROM PURE ORDER BY TRCID ASC LIMIT 5495,400"
    # query = "SELECT TRCID, CASRN, JSON FROM PURE WHERE TRCID IN (1143, 5523, 5822,5777, 5791)"    
    # query = "SELECT TRCID, CASRN, JSON FROM PURE WHERE TRCID IN (5523, 5822)"
    # query = "SELECT TRCID, CASRN, JSON FROM PURE WHERE TRCID IN (5791)"
    
    query = "SELECT TRCID, CASRN, JSON FROM PURE ORDER BY TRCID ASC"

    cursor.execute(query)

    total, saved, skipped = 0, 0, 0

    # Set up Excel workbook for unmatched components
    wb = Workbook()
    ws = wb.active
    ws.title = "Unmatched"
    ws.append(["Reason", "TRCID", "CASRN"])

    for row in cursor.fetchall():
        total += 1
        trcid, casrn, info = row

        # Skip if CASRN is missing
        casrn_str = str(casrn).strip()
        if not casrn_str:
            ws.append(["Missing CASRN", trcid, ""])
            skipped += 1
            continue

        # Validate JSON
        try:
            data = json.loads(info)
        except json.JSONDecodeError:
            ws.append(["Invalid JSON", trcid, casrn])
            skipped += 1
            continue

        # Save as JSON file
        file_name = f"{trcid}_CASNO_{casrn_str}.json"
        file_path = JSON_OUTPUT_DIR / file_name
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        print(f"Saved: {file_name}")
        saved += 1

    conn.close()

    # Save unmatched log if needed
    if skipped > 0:
        wb.save(UNMATCHED_FILE)
        print(f"\nUnmatched components written to: {UNMATCHED_FILE}")

    print("\nSummary:")
    print(f"Total records processed: {total}")
    print(f"Successfully extracted:  {saved}")
    print(f"Unmatched/skipped:      {skipped}")
    print("Done.")


def extract_smiles():
    print("Connecting to Compounds database...")
    if not DB_PATH_SMILES.exists():
        print(f"Database not found: {DB_PATH_SMILES}")
        return
    
    conn = sqlite3.connect(DB_PATH_SMILES)
    cursor = conn.cursor()

    query = "SELECT TRCID, CASRN, SMILES FROM COMPOUNDS ORDER BY TRCID ASC"
    cursor.execute(query)

    rows = cursor.fetchall()
    print(f"Total rows fetched: {len(rows)}")

    # Convert to DataFrame
    df = pd.DataFrame(rows, columns=["TRCID", "CASRN", "SMILES"])

    # Save to Excel (or CSV if preferred)
    df.to_excel(OUTPUT_FILE_SMILES, index=False)
    print(f"SMILES data exported to: {OUTPUT_FILE_SMILES}")

    conn.close()
    print("Done.")


if __name__ == "__main__":
    extract_and_log()
    extract_smiles()
    

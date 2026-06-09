"""
Script:
    35_clear_duplicate_casnum_from_xml.py

Purpose:
    This script clears duplicate or placeholder CAS numbers
    from selected XML component files.

Functionality:
    - Reads target XML file list from Excel
    - Processes generated XML component files
    - Identifies XML files requiring CAS modification
    - Detects placeholder CAS values such as 1003
    - Replaces selected casNum values with empty entries
    - Copies processed XML files into a new output folder
    - Generates CAS modification status reports

Input:
    - Excel file containing target XML file list
    - Generated XML component files

Output:
    - Updated XML component files with cleared CAS numbers
    - CAS modification Excel report
"""



import os
import shutil
from pathlib import Path
import pandas as pd
import re

from config import (
    RUN_YEAR,
    PREREQ_DIR,
    XML_DIR,
    XML_LIBRARY_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITES
# ==================================================

EXCEL_INPUTS_DIR = PREREQ_DIR / "excel_inputs"

EXCEL_INPUT = EXCEL_INPUTS_DIR / "8_NIST_Components_SameCAS.xlsx"

# ==================================================
# XML DIRECTORIES
# ==================================================

XML_SOURCE_DIR = XML_LIBRARY_DIR / "03_iccalc_updated_C1"

XML_OUTPUT_DIR = XML_LIBRARY_DIR / "04_casnum_updated"

REPORT_FILE = XML_LIBRARY_DIR / "04_casnum_update_report.xlsx"

#  Create only the new output directory
XML_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# VALIDATION 
# ==================================================

if not EXCEL_INPUT.exists():
    raise FileNotFoundError(f"Missing input: {EXCEL_INPUT}")

if not XML_SOURCE_DIR.exists():
    raise FileNotFoundError(f"Missing XML source directory: {XML_SOURCE_DIR}")

# ==================================================
# DEBUG (optional)
# ==================================================

print("Excel input:", EXCEL_INPUT)
print("XML source dir:", XML_SOURCE_DIR)
print("XML output dir:", XML_OUTPUT_DIR)
print("Report file:", REPORT_FILE)

# ---------------------------------------------------
# READ EXCEL FILE
# ---------------------------------------------------
df_input = pd.read_excel(EXCEL_INPUT)

if "XML_File" not in df_input.columns:
    raise ValueError("Column 'XML_File' not found in Excel file.")

target_files = set(df_input["XML_File"].dropna().astype(str).str.strip())

print(f"Total files in Excel list: {len(target_files)}")

# ---------------------------------------------------
# PROCESS XML FILES
# ---------------------------------------------------
results = []

for entry in os.scandir(XML_SOURCE_DIR):

    if not entry.is_file() or not entry.name.lower().endswith(".xml"):
        continue

    source_path = Path(entry.path)
    output_path = XML_OUTPUT_DIR / entry.name

    # Copy original by default
    shutil.copy2(source_path, output_path)

    modified = False

    if entry.name in target_files:
        with open(source_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Search for casNum="1003"
        if 'casNum="1003"' in content or 'casNum=" 1003"' in content or 'casNum="01003"' in content:
            new_content = content.replace('casNum="1003"', 'casNum=""')
            new_content = new_content.replace('casNum=" 1003"', 'casNum=""')
            new_content = new_content.replace('casNum="01003"', 'casNum=""')
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(new_content)

            modified = True

    results.append({
        "XML_File": entry.name,
        "In_Target_List": entry.name in target_files,
        "CAS_Modified": modified
    })

# ---------------------------------------------------
# GENERATE REPORT
# ---------------------------------------------------
df_report = pd.DataFrame(results)
df_report.to_excel(REPORT_FILE, index=False)

print("\nProcessing Completed")
print(f"Total XML processed: {len(results)}")
print(f"Modified files count: {df_report['CAS_Modified'].sum()}")
print(f"Updated XML folder created at: {XML_OUTPUT_DIR}")
print(f"Report generated at: {REPORT_FILE}")
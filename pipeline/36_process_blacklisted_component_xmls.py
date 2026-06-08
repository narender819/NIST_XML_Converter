"""
Script:
    36_process_blacklisted_component_xmls.py

Purpose:
    This script processes blacklisted component XML files
    using inclusion and exclusion rules defined in Excel sheets.

Functionality:
    - Reads blacklist inclusion and exclusion component lists
    - Processes generated XML component files
    - Skips excluded blacklisted components
    - Updates phase flags for included components
    - Removes vapor phase flags where required
    - Preserves unchanged XML files
    - Tracks missing XML files and processing errors
    - Generates blacklist processing status reports

Input:
    - Generated XML component files
    - Blacklist inclusion/exclusion Excel file

Output:
    - Updated blacklist-processed XML component files
    - Blacklist processing Excel report
"""



import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# PREREQUISITES
# ==================================================
PREREQ_DIR = (
    BASE_DIR
    / "prerequisites"
)

EXCEL_INPUTS_DIR = (
    PREREQ_DIR
    / "excel_inputs"
)

EXCEL_PATH = (
    EXCEL_INPUTS_DIR
    / "9_Components_Blacklist_Missing_Criticals.xlsx"
)

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_BASE_DIR = (
    BASE_DIR
    / "output"
    / RUN_YEAR
)

XML_BASE_DIR = (
    OUTPUT_BASE_DIR
    / "xml"
)

# ==================================================
# XML INPUT / OUTPUT
# ==================================================
XML_INPUT_DIR = (
    XML_BASE_DIR
    / "Libraryfiles_NIST"
    / "04_casnum_updated"
)

OUTPUT_DIR = (
    XML_BASE_DIR
    / "Libraryfiles_NIST"
    / "05_blacklist_processed"
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ---------------- READ TRCID COLUMN ----------------
def read_trcid_sheet(sheet_name):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()

    if "TRCID" not in df.columns:
        raise ValueError(f"TRCID column not found in {sheet_name}")

    trcids = (
        df["TRCID"]
        .dropna()
        .astype(int)
        .astype(str)
        .str.strip()
    )

    return set(trcids)

excluded_set = read_trcid_sheet("BlackList_Excluded")
included_set = read_trcid_sheet("BlackList_Included")

# ---------------- OVERLAP CHECK ----------------
overlap = excluded_set.intersection(included_set)
if overlap:
    print(f"⚠️ Warning: TRCID present in both sheets: {overlap}")

# ---------------- LOG TRACKING ----------------
report_data = []

excluded_skipped = []
included_updated = []
no_change = []
errors = []
processed_files = []

xml_files = list(XML_INPUT_DIR.glob("*.xml"))
xml_trcid_set = set()

# ---------------- PROCESS XML FILES ----------------
for xml_file in xml_files:

    filename = xml_file.name
    trcid = filename.replace("Comp-NIST-NST", "").replace(".xml", "").strip()
    xml_trcid_set.add(trcid)

    # -------- EXCLUDED --------
    if trcid in excluded_set:
        excluded_skipped.append(filename)
        report_data.append([trcid, filename, "Excluded", "", "", "Blacklisted"])
        continue

    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
    except Exception as e:
        errors.append((filename, str(e)))
        report_data.append([trcid, filename, "Error", "", "", str(e)])
        continue

    updated = False
    old_phase = ""
    new_phase = ""

    # -------- INCLUDED --------
    if trcid in included_set:
        for flags in root.findall(".//flags"):
            phase = flags.get("phase")

            if phase:
                old_phase = phase

            if phase and "V" in phase:
                new_phase = phase.replace("V", "").strip()

                if new_phase == "":
                    new_phase = "L"

                flags.set("phase", new_phase)
                updated = True

        if updated:
            included_updated.append(filename)
            report_data.append([trcid, filename, "Updated", old_phase, new_phase, "V removed"])
        else:
            no_change.append(filename)
            report_data.append([trcid, filename, "No Change", old_phase, old_phase, "No V present"])
    else:
        # Not in any list → just copy
        report_data.append([trcid, filename, "Unchanged", "", "", "Not in blacklist"])

    # -------- WRITE OUTPUT --------
    output_file = OUTPUT_DIR / filename
    try:
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
        processed_files.append(filename)
    except Exception as e:
        errors.append((filename, str(e)))
        report_data.append([trcid, filename, "Error", "", "", str(e)])

# ---------------- MISSING XML CHECK ----------------
missing_xml = excluded_set.union(included_set) - xml_trcid_set

for trcid in missing_xml:
    report_data.append([trcid, "", "Missing XML", "", "", "Not found in XML folder"])

# ---------------- SAVE REPORT ----------------
report_df = pd.DataFrame(report_data, columns=[
    "TRCID", "FileName", "Status", "OldPhase", "NewPhase", "Remarks"
])

report_path = (
    OUTPUT_DIR
    / "05_blacklist_processing_report.xlsx"
)
report_df.to_excel(report_path, index=False)

# ---------------- SUMMARY ----------------
print("\n================ SUMMARY REPORT ================\n")
print(f"Total XML files found     : {len(xml_files)}")
print(f"Excluded (skipped)        : {len(excluded_skipped)}")
print(f"Included (updated)        : {len(included_updated)}")
print(f"No change needed          : {len(no_change)}")
print(f"Successfully processed    : {len(processed_files)}")
print(f"Errors                    : {len(errors)}")
print(f"Missing XMLs              : {len(missing_xml)}")

print("\nReport saved at:", report_path)
print("\n================================================")
print("Processing completed successfully.\n")
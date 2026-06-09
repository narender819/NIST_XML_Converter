"""
Script:
    07_assign_dippr_family_groups.py

Purpose:
    Assign main component groups using DIPPR family and subfamily
    mappings.

Functionality:
    - Reads DIPPR family extraction data
    - Loads family and subfamily group mapping files
    - Maps family and subfamily groups
    - Assigns final main groups using matching logic
    - Detects and logs family/subfamily mismatches
    - Generates grouped component output files

Input:
    - DIPPR family extraction Excel file
    - DIPPR family/subfamily mapping Excel file

Output:
    - Excel file with assigned component groups
    - Mismatch log report
"""

import pandas as pd
from pathlib import Path

from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# PREREQUISITE DIRECTORIES (derived locally)
# ==================================================

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

EXCEL1 = PROCESSED_DIR / "5_DIPPR_Family_Extraction_withfamilyAbbrevations.xlsx"

EXCEL2 = EXCEL_INPUT_DIR / "4_DIPPR_Family_Mapping_with_Remarks.xlsx"

if not EXCEL1.exists():
    raise FileNotFoundError(f"Missing input: {EXCEL1}")

if not EXCEL2.exists():
    raise FileNotFoundError(f"Missing input: {EXCEL2}")

OUTPUT_MAIN = PROCESSED_DIR / "6_DIPPR_Family_Group_Assigned.xlsx"

OUTPUT_MISMATCH = PROCESSED_DIR / "6a_DIPPR_Group_Mismatch_Log.xlsx"



# ---------------------------------------------------
# NORMALIZATION FUNCTION
# ---------------------------------------------------
def normalize_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    if s.lower() in ["", "none", "nan"]:
        return None
    return s


# ---------------------------------------------------
# LOAD EXCEL FILES
# ---------------------------------------------------
print("Loading Excel files...")

df_main = pd.read_excel(EXCEL1)

df_family = pd.read_excel(EXCEL2, sheet_name="Family_Groups")
df_subfamily = pd.read_excel(EXCEL2, sheet_name="Subfamily_Groups")


# ---------------------------------------------------
# CLEAN + NORMALIZE FAMILY & SUBFAMILY LOOKUP TABLES
# ---------------------------------------------------
df_family["Code"] = df_family["Code"].astype(str).str.upper().str.strip()
df_family["Assigned_Main_Group"] = df_family["Assigned_Main_Group"].astype(str).str.strip()

df_subfamily["Dippr_Subfamily_Code"] = df_subfamily["Dippr_Subfamily_Code"].astype(str).str.lower().str.strip()
df_subfamily["Assigned_Group"] = df_subfamily["Assigned_Group"].astype(str).str.strip()

# Build lookup dictionaries
family_lookup = dict(zip(df_family["Code"], df_family["Assigned_Main_Group"]))
subfamily_lookup = dict(zip(df_subfamily["Dippr_Subfamily_Code"], df_subfamily["Assigned_Group"]))


# ---------------------------------------------------
# PROCESS COMPONENT DATA
# ---------------------------------------------------
family_groups = []
subfamily_groups = []
final_groups = []
conflict_flags = []
conflict_notes = []

mismatch_records = []


print("Processing DIPPR family/subfamily group assignment...")

for _, row in df_main.iterrows():

    # Normalize inputs
    fam_code = normalize_code(row.get("DIPPR_Family"))
    sub_code = normalize_code(row.get("DIPPR_Subfamily"))

    if fam_code:
        fam_code_u = fam_code.upper()
    else:
        fam_code_u = None

    if sub_code:
        sub_code_l = sub_code.lower()
    else:
        sub_code_l = None

    # Lookups
    fam_group = family_lookup.get(fam_code_u)
    sub_group = subfamily_lookup.get(sub_code_l)

    family_groups.append(fam_group if fam_group else "none")
    subfamily_groups.append(sub_group if sub_group else "none")

    # Decision Logic (FAMILY FIRST)
    if fam_group:
        final_group = fam_group
        note = "Family group used"
        conflict = False

        # Subfamily exists but differs -> log conflict
        if sub_group and sub_group != fam_group:
            conflict = True
            note = "Family/Subfamily mismatch"
            mismatch_records.append({
                "TRCID": row.get("TRCID"),
                "DIPPR_Family": fam_code_u,
                "Family_Group": fam_group,
                "DIPPR_Subfamily": sub_code_l,
                "Subfamily_Group": sub_group
            })

    elif sub_group:
        final_group = sub_group
        conflict = False
        note = "Family missing → using subfamily group"

    else:
        final_group = "none"
        conflict = False
        note = "No family/subfamily match found → Miscellaneous"

    final_groups.append(final_group if final_group else "none") 
    conflict_flags.append(conflict)
    conflict_notes.append(note)


# ---------------------------------------------------
# ADD NEW COLUMNS
# ---------------------------------------------------
df_main["Family_Assigned_Group"] = family_groups
df_main["Subfamily_Assigned_Group"] = subfamily_groups
df_main["Final_Main_Group"] = final_groups
# df_main["Group_Conflict_Flag"] = conflict_flags
# df_main["Group_Conflict_Note"] = conflict_notes


# ---------------------------------------------------
# SAVE MAIN OUTPUT
# ---------------------------------------------------
df_main.to_excel(OUTPUT_MAIN, index=False)
print(f"Main output saved {OUTPUT_MAIN}")



# ---------------------------------------------------
# SAVE MISMATCH LOG (if any)
# ---------------------------------------------------
if mismatch_records:
    df_mismatch = pd.DataFrame(mismatch_records)
    df_mismatch.to_excel(OUTPUT_MISMATCH, index=False)
    print(f"Mismatch log saved {OUTPUT_MISMATCH}")
else:
    print("No mismatches detected.")


print("Processing completed successfully.")

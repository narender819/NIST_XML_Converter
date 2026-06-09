"""
Script:
    10_assign_simsci_ids.py

Purpose:
    Assign existing or new SIMSCI IDs to components based on
    CAS numbers and component groups.

Functionality:
    - Reads component classification data
    - Loads master component list and library data
    - Matches components using CAS numbers
    - Assigns existing SIMSCI IDs where available
    - Generates new SIMSCI IDs for unmatched components
    - Applies group-wise SIMSCI ID assignment rules
    - Records SIMSCI ID source information
    - Generates final output with assigned SIMSCI IDs

Input:
    - Component classification Excel file
    - Master component list Excel file
    - Library extraction Excel file

Output:
    - Excel file with assigned SIMSCI IDs
"""

import pandas as pd
from collections import defaultdict
from openpyxl import load_workbook
from pathlib import Path

from utils import auto_adjust_column_widths
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

TARGET_FILE = PROCESSED_DIR / "8_DIPPR_Family_Extraction_All_With_Smiles_none_Predicted_Family.xlsx"

MASTER_FILE = EXCEL_INPUT_DIR / "5_Master_Component_List.xlsx"

LIBRARY_FILE = PROCESSED_DIR / "2_Libraries_XML_Component_Extract.xlsx"

OUTPUT_FILE = PROCESSED_DIR / "9_DIPPR_Family_Extraction_All_With_Smiles_none_Predicted_Family_SIMSCI_ASSIGNED.xlsx"

# ======================================================
# COLUMN DEFINITIONS
# ======================================================
TARGET_GROUP_COL = "Hybrid_Voted_Group"
MASTER_GROUP_COL = "Component Family"
LIB_GROUP_COL = "Component Family"

MASTER_SIMSCI_COL = "SimSci ID"
LIB_SIMSCI_COL = "SIMSCI ID"

# ======================================================
# HELPERS
# ======================================================
# def is_missing(val):
#     return val is None or (isinstance(val, str) and val.strip().lower() in {"", "nan", "none"})


def is_missing(val):
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and val.strip().lower() in {"", "nan", "none"}:
        return True
    return False



def norm(val):
    return str(val).strip()


def find_cas_column(df, candidates):
    for col in df.columns:
        c = col.strip().lower()
        for cand in candidates:
            if c == cand.lower():
                return col
    return None


# ======================================================
# LOAD FILES
# ======================================================
target_df = pd.read_excel(TARGET_FILE)
master_df = pd.read_excel(MASTER_FILE, header=5)
library_sheets = pd.read_excel(LIBRARY_FILE, sheet_name=None)

# ======================================================
# CLEAN COLUMN HEADERS
# ======================================================
target_df.columns = target_df.columns.str.strip()
master_df.columns = master_df.columns.str.strip()
for df in library_sheets.values():
    df.columns = df.columns.str.strip()

# ======================================================
# DETECT CAS COLUMNS
# ======================================================
TARGET_CAS_COL = find_cas_column(target_df, ["CASRN", "CAS", "CAS NO", "CASNO"])
MASTER_CAS_COL = find_cas_column(master_df, ["CAS NUMBER", "CASNO", "CAS", "CASRN"])

LIB_CAS_COL = None
for df in library_sheets.values():
    LIB_CAS_COL = find_cas_column(df, ["CASNO", "CAS NUMBER", "CAS", "CASRN"])
    if LIB_CAS_COL:
        break

if not TARGET_CAS_COL or not MASTER_CAS_COL or not LIB_CAS_COL:
    raise KeyError("CAS column not found in one or more input files")

# ======================================================
# NORMALIZE CAS
# ======================================================
target_df[TARGET_CAS_COL] = target_df[TARGET_CAS_COL].astype(str).str.strip()
master_df[MASTER_CAS_COL] = master_df[MASTER_CAS_COL].astype(str).str.strip()

for df in library_sheets.values():
    if LIB_CAS_COL in df.columns:
        df[LIB_CAS_COL] = df[LIB_CAS_COL].astype(str).str.strip()

# ======================================================
# BUILD CAS → SIMSCI LOOKUP
# ======================================================
cas_to_simsci = {}

# MASTER (priority)
for _, row in master_df.iterrows():
    cas = row.get(MASTER_CAS_COL)
    sim = row.get(MASTER_SIMSCI_COL)
    # if not is_missing(cas) and not is_missing(sim):
    #     cas_to_simsci[norm(cas)] = int(sim)
    if is_missing(cas) or is_missing(sim):
        continue

    try:
        cas_to_simsci[norm(cas)] = int(sim)
    except (ValueError, TypeError):
        continue


# LIBRARY
for sheet_df in library_sheets.values():
    for _, row in sheet_df.iterrows():
        cas = row.get(LIB_CAS_COL)
        sim = row.get(LIB_SIMSCI_COL)
        # if not is_missing(cas) and not is_missing(sim):
        #     cas_to_simsci.setdefault(norm(cas), int(sim))
        if is_missing(cas) or is_missing(sim):
            continue
        try:
            cas_to_simsci.setdefault(norm(cas), int(sim))
        except (ValueError, TypeError):
            continue


# ======================================================
# BUILD GROUP → MAX SIMSCI MAP
# ======================================================
group_max_id = defaultdict(int)

def update_group_max(df, group_col, sim_col):
    if group_col not in df.columns or sim_col not in df.columns:
        return
    for _, row in df.iterrows():
        g = row.get(group_col)
        s = row.get(sim_col)
        if not is_missing(g) and not is_missing(s):
            group_max_id[g] = max(group_max_id[g], int(s))


update_group_max(master_df, MASTER_GROUP_COL, MASTER_SIMSCI_COL)

for sheet_df in library_sheets.values():
    update_group_max(sheet_df, LIB_GROUP_COL, LIB_SIMSCI_COL)

print("Loaded group-wise max SIMSCI IDs:")
for g, v in sorted(group_max_id.items()):
    print(f"{g:20s} => {v}")

# ======================================================
# ASSIGN SIMSCI IDs (ONE PER ROW — GUARANTEED)
# ======================================================
assigned_ids = set()
simsci_ids = []
simsci_sources = []

for _, row in target_df.iterrows():
    cas = norm(row[TARGET_CAS_COL])
    group = row.get(TARGET_GROUP_COL)
    
    dippr_family = row.get("DIPPR_Family")

    # Normalize
    dippr_family = dippr_family.strip() if isinstance(dippr_family, str) else dippr_family

    # Default
    if is_missing(group):
        group = "Miscellaneous"

    # --- NEW RULES ---

    # Rule A: PO + Peroxides → Miscellaneous
    if dippr_family in {"PO", "none"} and group == "Peroxides":
        group = "Miscellaneous"


    # Rule B: none + Alkynes → Olefins
    if is_missing(dippr_family) and group == "Alkynes":
        group = "Olefins"


    # CAS MATCH
    if cas in cas_to_simsci:
        sim_id = cas_to_simsci[cas]
        source = "CAS_MATCH"

    # GENERATE NEW (GROUP-WISE)
    else:
        next_id = group_max_id[group] + 1

        while next_id in assigned_ids or next_id in cas_to_simsci.values():
            next_id += 1

        sim_id = next_id
        group_max_id[group] = next_id
        assigned_ids.add(sim_id)
        source = "GENERATED"

    simsci_ids.append(sim_id)
    simsci_sources.append(source)

# ======================================================
# FINAL SAFETY CHECK
# ======================================================
assert len(simsci_ids) == len(target_df)
assert len(simsci_sources) == len(target_df)

# ======================================================
# WRITE OUTPUT
# ======================================================
target_df["SIMSCI_ID"] = simsci_ids
target_df["SIMSCI_Source"] = simsci_sources

target_df.to_excel(OUTPUT_FILE, index=False)
auto_adjust_column_widths(OUTPUT_FILE)

print("\nSIMSCI ID assignment completed successfully")
print(f"Output written to:\n{OUTPUT_FILE}")

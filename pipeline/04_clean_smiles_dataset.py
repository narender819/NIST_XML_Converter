
"""
Script:
   04_clean_smiles_dataset.py

Purpose:
    This script cleans the extracted SMILES Excel data by standardizing
    CASRN values and removing blank or invalid entries.

Functionality:
    - Reads the extracted SMILES Excel file
    - Cleans CASRN values by:
        * Removing unwanted '.0' suffixes
        * Stripping extra spaces
        * Handling null values
    - Removes rows with missing or blank CASRN values
    - Generates a cleaned SMILES Excel report

Input:
    SMILES Excel file generated from the extraction process

Output:
    Cleaned Excel file containing valid CASRN and SMILES data
"""


import os
import pandas as pd

from pathlib import Path

from config import (
    RUN_YEAR,
    OUTPUT_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================

SMILES_OUTPUT_DIR = OUTPUT_DIR / "smiles"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

INPUT_FILE = SMILES_OUTPUT_DIR / f"1_compounds_smiles_{RUN_YEAR}.xlsx"

if not INPUT_FILE.exists():
    raise FileNotFoundError(f"Missing input: {INPUT_FILE}")

OUTPUT_FILE = SMILES_OUTPUT_DIR / f"2_compounds_smiles_{RUN_YEAR}_removedblanks.xlsx"

def clean_cas(cas):
    """Clean CASRN: remove .0 if float-like, strip spaces, drop NaN."""
    if pd.isna(cas):
        return None
    cas_str = str(cas).strip()
    if cas_str.endswith(".0"):  # e.g., '71432.0' → '71432'
        cas_str = cas_str[:-2]
    return cas_str if cas_str else None

# Load Excel
df = pd.read_excel(INPUT_FILE)

# Clean CASRN
df['CASRN'] = df['CASRN'].apply(clean_cas)

# Drop rows where CASRN is empty
df = df.dropna(subset=['CASRN'])
df = df[df['CASRN'] != ""]

# Save cleaned file
df.to_excel(OUTPUT_FILE, index=False)

print(f"Cleaned file saved as: {OUTPUT_FILE}")
print(f"Original rows: {len(pd.read_excel(INPUT_FILE))}, After cleaning: {len(df)}")

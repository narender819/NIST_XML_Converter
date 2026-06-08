"""
Script:
    27_extract_nist_scp_correlation_coefficients.py

Purpose:
    This script extracts Solid Enthalpy (SCP) correlation
    coefficients and metadata from generated NIST XML files.

Functionality:
    - Reads generated NIST XML component files
    - Extracts component identification details
    - Extracts SolidEnthalpy correlation information
    - Extracts SCP equation numbers and temperature ranges
    - Extracts and standardizes SCP correlation coefficients
    - Performs unit conversion when required
    - Generates a consolidated SCP coefficient Excel report

Input:
    Generated NIST XML component files

Output:
    Excel report containing NIST SCP correlation coefficients and metadata
"""

import pandas as pd
import xml.etree.ElementTree as ET
from pathlib import Path

def normalize_cas(cas):
    """
    Normalize CAS to digit-only string: '71-43-2' or '71432' → '71432'
    """
    if cas is None:
        return None
    s = str(cas).strip()
    if not s:
        return None
    # Remove hyphens and keep only digits
    s = s.replace("-", "").replace(" ", "")
    s = "".join(ch for ch in s if ch.isdigit())
    return s if s else None

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = (
    BASE_DIR
    / "output"
    / RUN_YEAR
)

PROCESSED_DIR = (
    OUTPUT_DIR
    / "processed"
    / "full_library"
)

XML_DIR = (
    OUTPUT_DIR
    / "xml"
    / "Libraryfiles_NIST"
    / "01_generated"
)

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================
OUTPUT_EXCEL = (
    PROCESSED_DIR
    / f"19_NIST_SIMSCI_SCP_Key.xlsx"
)

PROCESSED_DIR.mkdir(
    parents=True,
    exist_ok=True
)
MAX_COEFFS = 8

rows = []

# ---------------------------------------------------
# LOOP XML FILES
# ---------------------------------------------------
for xml_file in XML_DIR.glob("*.xml"):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # -----------------------------
        # Extract ID block
        # -----------------------------
        id_tag = root.find(".//id")
        if id_tag is None:
            continue

        trcid = id_tag.get("refCode")
        comp_name = id_tag.get("name")
        cas_raw = id_tag.get("casNum")
        alias = id_tag.text.strip() if id_tag.text else None

        # Normalize CAS
        cas = normalize_cas(cas_raw)
        if cas is None:
            continue  # Skip if no valid CAS

        # -----------------------------
        # Find SolidEnthalpy correlation
        # -----------------------------
        corr = root.find(".//correlation[@name='SolidEnthalpy']")
        if corr is None:
            continue

        eq_no = corr.get("equation")
        tmin = corr.get("tMin")
        tmax = corr.get("tMax")
        prop_unit = corr.get("propUnit")

        # -----------------------------
        # Extract coefficients
        # -----------------------------
        coeff_text = corr.text.strip()
        coeffs = [float(x) for x in coeff_text.split()]

        # -----------------------------
        # Unit conversion (J/g-mole → J/kg-mole)
        # -----------------------------
        if prop_unit == "J/g-mole":
            coeffs = [c * 1000.0 for c in coeffs]
            prop_unit = "J/kg-mole"

        # Pad coefficients
        coeff_padded = coeffs + [None] * (MAX_COEFFS - len(coeffs))

        row = {
            "TRCID": int(trcid),
            "CASNO": cas,  # Now normalized: '71432'
            "ComponentName": comp_name,
            "Aliases": alias,
            "SCPEqn": int(eq_no) if eq_no else None,
            "SCPtmin (K)": float(tmin) if tmin else None,
            "SCPtmax (K)": float(tmax) if tmax else None,
            "PropUnit": prop_unit
        }

        # Add coefficients columns C1 → Cn
        for i in range(MAX_COEFFS):
            row[f"C{i+1}"] = coeff_padded[i]

        rows.append(row)

    except Exception as e:
        print(f"Error processing {xml_file.name}: {e}")

# ---------------------------------------------------
# Create DataFrame
# ---------------------------------------------------
df = pd.DataFrame(rows)
df.sort_values(by="TRCID", inplace=True)

# ---------------------------------------------------
# Save NIST sheet
# ---------------------------------------------------
with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name="NIST_SCP_Key", index=False)

print("NIST_SCP_Key extraction completed successfully.")
print(f"Total NIST SCP components: {len(df)}")
print(f"Sample CAS: {df['CASNO'].head().tolist()}")  # Should show 71432 format

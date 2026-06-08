"""
Script:
   01_extract_library_xml_component_data.py

Purpose:
    Extract component metadata and enthalpy correlation temperature ranges
    from SIMSCI XML library files and generate a consolidated Excel report.

Input:
    XML library files from BASE_PATH

Output:
    Multi-sheet Excel workbook containing extracted component data
"""


from logging import root
import os
import xml.etree.ElementTree as ET
import pandas as pd

from pathlib import Path

ROOT_DIR = Path(r"D:\NIST_XML_Converter")

PREREQ_DIR = ROOT_DIR / "prerequisites"

LIBRARIES_DIR = PREREQ_DIR / "libraries"

OUTPUT_2025_DIR = ROOT_DIR / "output" / "2025"

PROCESSED_DIR = OUTPUT_2025_DIR / "processed" / "full_library"

XML_OUTPUT_DIR = OUTPUT_2025_DIR / "xml"


# ---------------- CONFIG ----------------
BASE_PATH = LIBRARIES_DIR / "SIMSCI" / "Libraries_xmlfiles_2025"
OUTPUT_EXCEL = PROCESSED_DIR / "1_Libraries_XML_Component_Extract_nodipprsponsor.xlsx"




# SIMSCI Family Mapping
FAMILY_MAP = {
    1: "Acids",
    2: "Alcohols",
    3: "Aldehydes",
    4: "Amides",
    5: "Amines",    
    6: "Ethers",
    7: "Esters",
    8: "Halogenated compounds",
    9: "Paraffins",
    10: "Naphthenes",
    11: "Olefins",
    12: "Aromatics",
    13: "Ketones",
    14: "Nitrogen compounds",
    15: "Sulfur compounds",
    16: "Miscellaneous",
    17: "Silicon compounds",
    18: "Salts",
    19: "Elements",
}

# ---------------- FUNCTIONS ----------------
def extract_component_family(simsci_id):
    """
    Extract component family based on SIMSCI ID length.
    """
    if not simsci_id or not simsci_id.isdigit():
        return "Unknown"

    if len(simsci_id) == 7:
        key = int(simsci_id[0])
    elif len(simsci_id) == 8:
        key = int(simsci_id[:2])
    else:
        return "Unknown"

    return FAMILY_MAP.get(key, "Unknown")

def extract_library_id(id_elem):
    """
    Extract first alias inside <id> tag text.
    Example: '1C2H 1C2MCH CMECHEXN' -> '1C2H'
    """
    if id_elem is None:
        return None

    if id_elem.text is None:
        return None

    text = id_elem.text.strip()
    if not text:
        return None

    return text.split()[0]


def extract_tmin_tmax(parent, corr_name):
    """
    Extract tMin / tMax from SIMSCI <correlation> blocks.
    Safe: returns (None, None) if correlation not found.
    """
    for corr in parent.findall(".//correlation"):
        if corr.attrib.get("name") == corr_name:
            tmin = corr.attrib.get("tMin")
            tmax = corr.attrib.get("tMax")

            try:
                tmin = float(tmin) if tmin is not None else None
            except ValueError:
                tmin = None

            try:
                tmax = float(tmax) if tmax is not None else None
            except ValueError:
                tmax = None

            return tmin, tmax

    return None, None


def parse_xml_file(xml_file: str) -> list:
    rows = []
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for comp in root.iter("comp"):

            id_elem = comp.find("id")
            if id_elem is None:
                continue

            simsci_id = id_elem.attrib.get("number") or None
            name = id_elem.attrib.get("name") or None
            formula = id_elem.attrib.get("formula") or None
            casno = id_elem.attrib.get("casNum") or None
            library_id = extract_library_id(id_elem)

            # --- NEW: Enthalpy validity ranges (SAFE ADDITION) ---
            LEtmin, LEtmax = extract_tmin_tmax(comp, "LiquidEnthalpy")
            IEtmin, IEtmax = extract_tmin_tmax(comp, "IdealEnthalpy")
            SEtmin, SEtmax = extract_tmin_tmax(comp, "SolidEnthalpy")

            rows.append({
                # --- Existing columns (UNCHANGED) ---
                "SIMSCI ID": simsci_id,
                "LibraryID": library_id,
                "CASNO": casno,
                "Name": name,
                "Formula": formula,
                "Component Family": extract_component_family(simsci_id),

                # --- New columns (APPENDED, non-breaking) ---
                "LEtmin": LEtmin,
                "LEtmax": LEtmax,
                "IEtmin": IEtmin,
                "IEtmax": IEtmax,
                "SEtmin": SEtmin,
                "SEtmax": SEtmax,
            })

    except Exception as e:
        print(f"[ERROR] Failed parsing {xml_file}: {e}")

    return rows




# ---------------- MAIN WORKFLOW ----------------
def generate_excel_from_xml():
    writer = pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="w")
    for root, dirs, files in os.walk(BASE_PATH):
        # Skip BASE_PATH root itself
        if root == BASE_PATH:
            continue

        # Build sheet name
        rel_path = os.path.relpath(root, BASE_PATH)
        sheet_name = rel_path.replace(os.sep, "_")[:31]  # Excel limit

        all_rows = []

        for file in files:
            # if file.lower().startswith("comp") and file.lower().endswith(".xml"):
            if (file.lower().endswith(".xml") and not (file.startswith("__") or file.startswith("Bank"))):
                xml_file_path = os.path.join(root, file)
                all_rows.extend(parse_xml_file(xml_file_path))

        if all_rows:
            df = pd.DataFrame(all_rows)
            df.insert(0, "S.No", range(1, len(df) + 1))
            df_out = df.fillna("NA")

            df_out.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Processed sheet: {sheet_name} with {len(df)} records.")

    writer.close()
    print(f"Excel file generated successfully: {OUTPUT_EXCEL}")


# ---------------- EXECUTION ----------------
if __name__ == "__main__":
    generate_excel_from_xml()

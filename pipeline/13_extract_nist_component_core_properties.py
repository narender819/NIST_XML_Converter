"""
Script:
    13_extract_nist_component_core_properties.py

Purpose:
    This script extracts key thermodynamic and temperature-dependent
    component properties from processed NIST JSON files and generates
    a consolidated Excel report.

Functionality:
    - Reads processed component JSON files
    - Extracts core thermodynamic properties
    - Extracts temperature-dependent property limits and equations
    - Calculates derived reference temperatures and enthalpy values
    - Normalizes and validates CAS numbers
    - Handles missing values using sentinel defaults
    - Consolidates all extracted data into a sorted Excel report

Input:
    Processed component JSON files

Output:
    Excel report containing extracted core thermodynamic properties
"""
import json
from unittest import result
import pandas as pd
from pathlib import Path
import numpy as np
import re


from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# CONSTANTS
# ==================================================

SENTINEL = -9999

# ==================================================
# INPUT FOLDERS
# ==================================================

COMPONENTS_MASTER_DIR = PROCESSED_DIR / "1_components_Inmaster_withsimsciid"
COMPONENTS_ASSIGNED_DIR = PROCESSED_DIR / "3_components_notInmaster_assignedsimsciid"

#  Ensure these exist (important here)
COMPONENTS_MASTER_DIR.mkdir(parents=True, exist_ok=True)
COMPONENTS_ASSIGNED_DIR.mkdir(parents=True, exist_ok=True)

folders = [
    COMPONENTS_MASTER_DIR,
    COMPONENTS_ASSIGNED_DIR
]

# ==================================================
# OUTPUT FILE
# ==================================================

output_excel = PROCESSED_DIR / f"10_NIST_Component_KeyThermoProperties_{RUN_YEAR}.xlsx"

rows = []


def extract_enthalpy_limits(temp_deps):
    """Extract tmin/tmax/eq from first enthalpy entries; NA if missing."""
    result = {
        'LCPtmin (K)': 'NA', 'LCPtmax (K)': 'NA', 'LCPEqn': 'NA',
        'SCPtmin (K)': 'NA', 'SCPtmax (K)': 'NA', 'SCPEqn': 'NA',
        'ICPtmin (K)': 'NA', 'ICPtmax (K)': 'NA', 'ICPEqn': 'NA',
        'HVAPtmin (K)': 'NA', 'HVAPtmax (K)': 'NA', 'HVAPEqn': 'NA',
        'STtmin (K)': 'NA', 'STtmax (K)': 'NA', 'STEqn': 'NA',
        'LDtmin (K)': 'NA', 'LDtmax (K)': 'NA', 'LDEqn': 'NA',
        'VPtmin (K)': 'NA', 'VPtmax (K)': 'NA', 'VPEqn': 'NA'
    }
    for prop_name, entries in temp_deps.items():
        if not entries:
            continue
        entry = entries[0]
        tmin = entry.get('tmin')
        tmax = entry.get('tmax')
        eqn = entry.get('equation')
        if prop_name == 'LiquidEnthalpy':
            result['LCPtmin (K)'] = tmin if tmin is not None else 'NA'
            result['LCPtmax (K)'] = tmax if tmax is not None else 'NA'
            result['LCPEqn'] = eqn if eqn is not None else 'NA'
        elif prop_name == 'SolidEnthalpy':
            result['SCPtmin (K)'] = tmin if tmin is not None else 'NA'
            result['SCPtmax (K)'] = tmax if tmax is not None else 'NA'
            result['SCPEqn'] = eqn if eqn is not None else 'NA'
        elif prop_name == 'IdealEnthalpy':
            result['ICPtmin (K)'] = tmin if tmin is not None else 'NA'
            result['ICPtmax (K)'] = tmax if tmax is not None else 'NA'
            result['ICPEqn'] = eqn if eqn is not None else 'NA'

            # -------------------------------
            # NEW: ICP reference temperature (Tnbp)
            # -------------------------------
            nbp = result.get("NBP (K)")
            if nbp not in (None, "NA"):
                result["ICP_Tnbp (K)"] = nbp
            elif tmin is not None and tmax is not None:
                # ensure numeric
                try:
                    tmin_f = float(tmin)
                    tmax_f = float(tmax)
                    result["ICP_Tnbp (K)"] = tmin_f + 0.05 * (tmax_f - tmin_f)
                except (TypeError, ValueError):
                    result["ICP_Tnbp (K)"] = "NA"
            else:
                result["ICP_Tnbp (K)"] = "NA"

        elif prop_name == 'LatentHeat':
            result['HVAPtmin (K)'] = tmin if tmin is not None else 'NA'
            result['HVAPtmax (K)'] = tmax if tmax is not None else 'NA'
            result['HVAPEqn'] = eqn if eqn is not None else 'NA'

            # NEW: midpoint temperature for SIMSCI ICP protocol fallback
            if tmin is not None and tmax is not None:
                try:
                    tmin_f = float(tmin)
                    tmax_f = float(tmax)
                    result['HVAP_Tmid (K)'] = (tmin_f + tmax_f) / 2.0
                except (TypeError, ValueError):
                    result['HVAP_Tmid (K)'] = 0
            else:
                result['HVAP_Tmid (K)'] = 0
        elif prop_name == 'SurfaceTension':
            result['STtmin (K)'] = tmin if tmin is not None else 'NA'
            result['STtmax (K)'] = tmax if tmax is not None else 'NA'
            result['STEqn'] = eqn if eqn is not None else 'NA'
        
        elif prop_name == 'LiquidDensity':
            result['LDtmin (K)'] = tmin if tmin is not None else 'NA'
            result['LDtmax (K)'] = tmax if tmax is not None else 'NA'
            result['LDEqn'] = eqn if eqn is not None else 'NA'
        elif prop_name == 'VaporPressure':
            result['VPtmin (K)'] = tmin if tmin is not None else 'NA'
            result['VPtmax (K)'] = tmax if tmax is not None else 'NA'
            result['VPEqn'] = eqn if eqn is not None else 'NA'

    return result


# -----------------------------
# Helper: parse filename
# -----------------------------
def parse_filename(filename):
    stem = Path(filename).stem
    parts = stem.split("_CASNO_")
    trcid = parts[0]
    casrn = parts[1] if len(parts) > 1 else None
    return trcid, casrn

# ------------------------------
# hhfusion calculation - FIXED
# ------------------------------
def safe_hfusion_to_jmol(hfusion_raw):
    """Convert HFUSION to J/mol, handling sentinel/NA safely"""
    if hfusion_raw in (None, SENTINEL, "NA"):
        return 0.0
    
    try:
        hfusion_float = float(hfusion_raw)
        if hfusion_float == SENTINEL:
            return 0.0
        return hfusion_float * 1000.0  # kJ/mol → J/mol
    except (TypeError, ValueError):
        return 0.0


# -----------------------------
# Read JSON files
# -----------------------------
for folder in folders:
    for json_file in Path(folder).glob("*.json"):
        try:
            trcid, casrn = parse_filename(json_file.name)

            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            props = data.get("properties", {})
            temp_deps = data.get("temp-dep_properties", {})
            enthalpy_limits = extract_enthalpy_limits(temp_deps)

            # ------------------------------
            # Core properties
            # ------------------------------
            NMP = props.get("NMP", SENTINEL)
            MW = props.get("MW", SENTINEL)
            ACENTRIC = props.get("ACENTRIC", SENTINEL)
            

            # ------------------------------
            # SCP_Tnmp calculation
            # ------------------------------
            SCPtmin = enthalpy_limits.get("SCPtmin (K)", SENTINEL)
            SCPtmax = enthalpy_limits.get("SCPtmax (K)", SENTINEL)

            # Normalize possible 'NA'/string values to numeric or sentinel
            def safe_float_or_sentinel(v):
                if v in (None, "NA"):
                    return SENTINEL
                try:
                    return float(v)
                except (TypeError, ValueError):
                    return SENTINEL

            SCPtmin_val = safe_float_or_sentinel(SCPtmin)
            SCPtmax_val = safe_float_or_sentinel(SCPtmax)

            # ACENTRIC might also be string; normalize
            if ACENTRIC not in (None, SENTINEL, "NA"):
                try:
                    ACENTRIC_val = float(ACENTRIC)
                except (TypeError, ValueError):
                    ACENTRIC_val = SENTINEL
            else:
                ACENTRIC_val = SENTINEL

            # MW might also be string; normalize
            if MW not in (None, SENTINEL, "NA"):
                try:
                    MW_val = float(MW)
                except (TypeError, ValueError):
                    MW_val = SENTINEL
            else:
                MW_val = SENTINEL

            # NMP might also be string; normalize
            if NMP not in (None, SENTINEL, "NA"):
                try:
                    NMP_val = float(NMP)
                except (TypeError, ValueError):
                    NMP_val = SENTINEL
            else:
                NMP_val = SENTINEL

            if NMP_val != SENTINEL:
                SCP_Tnmp = NMP_val
            elif SCPtmin_val != SENTINEL and SCPtmax_val != SENTINEL:
                SCP_Tnmp = SCPtmin_val + 0.95 * (SCPtmax_val - SCPtmin_val)
            else:
                SCP_Tnmp = SENTINEL

 
            rows.append({
                "TRCID": trcid,
                "CASNO": casrn or data.get("CASNO"),
                "ComponentName": data.get("name"),
                "Aliases": ", ".join(data.get("aliases", [])),
                "NMP": props.get("NMP"),
                "MW": props.get("MW"),
                "ACENTRIC": props.get("ACENTRIC"),
                "NBP": props.get("NBP"),
                "TC": props.get("TC"),
                "PC": props.get("PC"),
                "TTP": props.get("TTP"),
                "HFUSIONNMP": props.get("HFUSIONNMP"),
                **enthalpy_limits,
                "SCP_Tnmp (K)": SCP_Tnmp,
                "hhfusion (J/kg-mole)":safe_hfusion_to_jmol(props.get("HFUSIONNMP"))
            })

        except Exception as e:
            print(f"Error processing {json_file}: {e}")




# -----------------------------
# Create DataFrame
# -----------------------------
df = pd.DataFrame(rows)

# -----------------------------
# Rename columns to include units
# -----------------------------
df.rename(columns={
    "NMP": "NMP (K)",
    "NBP": "NBP (K)",
    "TC": "TC (K)",
    "PC": "PC (kPa)",
    "TTP": "TTP (K)",
    "HFUSIONNMP": "HFUSIONNMP (kJ/kg-mol)",
}, inplace=True)

CAS_PATTERN = re.compile(r"^\d{1,7}-\d{2}-\d$")


def normalize_cas(cas):
    if pd.isna(cas):
        return None
    return str(cas).strip()


def is_valid_cas(cas):
    if not cas:
        return False
    return bool(CAS_PATTERN.match(cas))


def auto_format_cas(cas):
    if pd.isna(cas):
        return None

    cas_str = str(cas).strip()

    # already formatted
    if "-" in cas_str:
        return cas_str

    # digits-only CAS → XXXXXXX → XXXX-XX-X
    if cas_str.isdigit() and len(cas_str) >= 4:
        return f"{cas_str[:-3]}-{cas_str[-3:-1]}-{cas_str[-1]}"

    return cas_str


# -----------------------------
# Fill numeric missing values with sentinel
# -----------------------------
NUMERIC_SENTINEL = -9999

numeric_columns = [
    "NMP (K)",
    "NBP (K)",
    "TC (K)",
    "PC (kPa)",
    "HFUSIONNMP (kJ/kg-mol)",
    "HVAPtmin (K)",
    "HVAPtmax (K)",
    "HVAP_Tmid (K)",
    "ICP_Tnbp (K)",
    "SCP_Tnmp (K)",  
    "MW",   
    "ACENTRIC" # include newly calculated column
]

df[numeric_columns] = df[numeric_columns].apply(
    pd.to_numeric, errors="coerce"
).fillna(NUMERIC_SENTINEL)

# -----------------------------
# Fill text columns
# -----------------------------
text_columns = [
    "TRCID",
    "CASNO",
    "ComponentName",
    "Aliases"
]

df[text_columns] = df[text_columns].fillna("NA")

# Explicitly cast TRCID to numeric
df["TRCID"] = pd.to_numeric(df["TRCID"], errors="raise")

# Sort by TRCID
df.sort_values(by="TRCID", ascending=True, inplace=True)

# Normalize + auto-format CAS
df["CASNO"] = df["CASNO"].apply(normalize_cas).apply(auto_format_cas)

# Validate CAS format
df["CAS_Valid"] = df["CASNO"].apply(is_valid_cas)

# Write to Excel
df.to_excel(output_excel, index=False)
print(f"Sorted Excel generated successfully: {output_excel}")

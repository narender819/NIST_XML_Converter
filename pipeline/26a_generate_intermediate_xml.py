"""
Script:
    26_generate_simsci_xml_from_processed_json.py

Purpose:
    This script converts processed component JSON files into
    SimSci-compatible XML component files.

Functionality:
    - Reads processed component JSON files
    - Loads property, unit, and equation mapping configurations
    - Extracts fixed and temperature-dependent properties
    - Maps property equations and units
    - Processes phase applicability information
    - Applies correlation parameter transformations
    - Generates SimSci-compatible XML component structures
    - Formats and writes XML component files
    - Generates XML files for both in-master and new components

Input:
    - Processed component JSON files
    - Property mapping configuration Excel file

Output:
    SimSci-compatible XML component files
"""
INVALID_CAS = []

import os
import json
import pandas as pd
import xml.etree.ElementTree as ET
import xml.dom.minidom
import re
import logging
import math

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

from pathlib import Path

from config import (
    RUN_YEAR,
    PREREQ_DIR,
    PROCESSED_DIR,
    XML_DIR,
    XML_LIBRARY_DIR,
    INTERMEDIATE_XML_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# CONSTANTS
# ==================================================

REF_CODE_DEFAULT = "11005"

# ==================================================
# PREREQUISITES
# ==================================================

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

CONFIG_FILE = EXCEL_INPUT_DIR / "6_NIST_Property_Mappings_Template.xlsx"

# ==================================================
# INPUT DIRECTORIES
# ==================================================

MASTER_DIR = PROCESSED_DIR / "1_components_Inmaster_withsimsciid"

ASSIGNED_DIR = PROCESSED_DIR / "3_components_notInmaster_assignedsimsciid"

# ==================================================
# OUTPUT DIRECTORY
# ==================================================

XML_OUTPUT_DIR = INTERMEDIATE_XML_DIR

XML_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# VALIDATION 
# ==================================================

if not CONFIG_FILE.exists():
    raise FileNotFoundError(f"Missing input: {CONFIG_FILE}")

if not MASTER_DIR.exists():
    raise FileNotFoundError(f"Missing folder: {MASTER_DIR}")

if not ASSIGNED_DIR.exists():
    raise FileNotFoundError(f"Missing folder: {ASSIGNED_DIR}")

# ==================================================
# DEBUG (optional)
# ==================================================

print("Config file:", CONFIG_FILE)
print("Master dir:", MASTER_DIR)
print("Assigned dir:", ASSIGNED_DIR)
print("XML output dir:", XML_OUTPUT_DIR)

# -------------------------
# Utility: CAS normalization
# -------------------------
def normalize_cas(cas):
    """Normalize CAS number while keeping hyphens in standard CAS format (xxxxx-yy-z)."""
    if cas is None:
        return None
    cas_str = str(cas).strip()
    if cas_str == "":
        return None
    # If already in standard format (e.g., 64-17-5), normalize leading zeros removal in first part
    m = re.match(r'^\s*0*(\d{1,7})-(\d{2})-(\d)\s*$', cas_str)
    if m:
        first = m.group(1) or '0'
        middle = m.group(2)
        last = m.group(3)
        return f"{first}-{middle}-{last}"
    # Otherwise, get digits and recompose
    digits = ''.join(ch for ch in cas_str if ch.isdigit())
    if len(digits) < 5:
        return None
    first = digits[:-3].lstrip('0') or '0'
    middle = digits[-3:-1]
    last = digits[-1]
    formatted = f"{first}-{middle}-{last}"
    if re.match(r'^\d{1,7}-\d{2}-\d$', formatted):
        return formatted
    return None
def validate_cas_checksum(cas):

    if not cas: return False

    digits = str(cas).replace("-", "")

    check_digit = int(digits[-1])

    total = sum((i+1)*int(d) for i,d in enumerate(digits[:-1][::-1]))

    return total % 10 == check_digit
def is_valid_cas(cas):
    if not cas or not isinstance(cas, str):
        return False
    return re.match(r'^\d{1,7}-\d{2}-\d$', cas) is not None

# -------------------------
# Load config helpers (UPDATED - Added model mapping)
# -------------------------
def load_uom_mapping(config_path):
    df = pd.read_excel(config_path, sheet_name="FIXED_PROP_NAMES")
    uom_map = {}
    for _, row in df.iterrows():
        sim_name = str(row.get('SimSci_Name', '')).strip()
        sim_uom = str(row.get('SimSci_UOM', '')).strip() if pd.notna(row.get('SimSci_UOM')) else ''
        if sim_name:
            uom_map[sim_name] = sim_uom.lower()
    return uom_map

def load_tempdep_property_config(config_path):
    df = pd.read_excel(config_path, sheet_name="TDEP_COR_NAMES")
    prop_config = {}
    for _, row in df.iterrows():
        sim_name = str(row.get('SimSci_Name', '')).strip()
        if sim_name:
            prop_config[sim_name] = {
                'tempUnit': str(row.get('Temp_Unit', '') or row.get('TempUnit', '') or '').strip(),
                'propUnit': str(row.get('SimSci_UOM', '') or row.get('PropUnit', '') or '').strip(),
                'moleWtBasis': str(row.get('MoleWt_Basis', '') or row.get('MoleWtBasis', '') or '').strip(),
                'logBasis': str(row.get('Log_Basis', '') or row.get('Log_Basis', '') or '').strip(),
                'equation': str(row.get('Equation', '')).strip(),
                'Model_Name': str(row.get('Model_Name', '') or '').strip()
            }
    return prop_config

def load_model_equation_mapping(config_path):
    """NEW: Load model-to-equation mapping from MODEL_EQUATIONS sheet."""
    try:
        df = pd.read_excel(config_path, sheet_name="MODEL_EQUATIONS")
        mapping = {}
        for _, row in df.iterrows():
            model = str(row.get('Model_Name', '')).strip().lower()
            eq = str(row.get('Default_Equation', '')).strip()
            prop = str(row.get('Property_Type', '')).strip().lower()
            if model and eq:
                mapping.setdefault(prop, {})[model] = eq
        logging.info("Loaded model-equation mapping")
        return mapping
    except FileNotFoundError:
        logging.warning("MODEL_EQUATIONS sheet not found - using existing logic only")
        return {}
    except Exception as e:
        logging.warning(f"Failed to load MODEL_EQUATIONS: {e}")
        return {}

# -------------------------
# Phase mapping
# -------------------------
PHASE_MAP = {
    "C1": "S", "S": "S", "L": "VL", "G": "VL", "IG": "VL",
    "VL": "VL", "LS": "LS", "VLS": "VLS",
}

def map_phases_to_flag(applicable_phases_str):
    if not applicable_phases_str:
        return "VLS"
    phases = [p.strip().upper() for p in applicable_phases_str.split(",")]
    normalized = set()
    for p in phases:
        mapped = PHASE_MAP.get(p)
        if mapped:
            normalized.add(mapped)
    # Priority collapse
    if "VLS" in normalized: return "VLS"
    if "LS" in normalized and "VL" in normalized: return "VLS"
    if "LS" in normalized: return "LS"
    if "VL" in normalized: return "VL"
    if "S" in normalized: return "S"
    return "VLS"

def json_key_to_xml_attr_name(key):
    if key.lower() == "applicable phases":
        return "ApplicablePhases"
    return key

def pretty_xml(element):
    """Return formatted XML string without extra XML declaration."""
    raw_str = ET.tostring(element, 'utf-8')
    reparsed = xml.dom.minidom.parseString(raw_str)
    pretty_str = reparsed.toprettyxml(indent=" ", encoding="utf-8").decode("utf-8")
    pretty_str = re.sub(r'<\?xml[^>]+\?>\s*', '', pretty_str)
    return pretty_str.strip()

# -------------------------
# NEW: Robust equation resolution (100% backward compatible)
# -------------------------
def resolve_equation(model, prop_name, equation_from_json, model_eq_map, tempdep_config):
    """Priority: JSON > Model map > Config match > Fallback"""
    
    # 1) Explicit JSON equation (highest priority)
    if equation_from_json and str(equation_from_json).strip().lower() not in ("", "nil"):
        return str(equation_from_json).strip()
    
    # 2) NEW: Model-to-equation mapping (fixes VDN PolynomialDensity)
    prop_key = prop_name.strip().lower()
    model_lower = str(model.get('name', '')).strip().lower()
    if prop_key in model_eq_map and model_lower in model_eq_map[prop_key]:
        logging.info(f"Using model map: {prop_name} '{model.get('name')}' → eq '{model_eq_map[prop_key][model_lower]}'")
        return model_eq_map[prop_key][model_lower]
    
    # 3) Existing config model name matching (BACKWARD COMPATIBLE)
    prop_key = next((k for k in tempdep_config.keys() if k.strip().lower() == prop_name.strip().lower()), None)
    if prop_key:
        cfg = tempdep_config[prop_key]
        config_model_names = [m.strip().lower() for m in str(cfg.get("Model_Name", "")).split("/") if m.strip() != ""]
        config_equations = [e.strip() for e in str(cfg.get("equation", "")).split()]
        
        matched_index = -1
        for i, mn in enumerate(config_model_names):
            if mn and mn in model_lower:
                matched_index = i
                break
        if matched_index >= 0 and matched_index < len(config_equations):
            eqv = config_equations[matched_index]
            if str(eqv).strip().lower() != "nil":
                return str(eqv).strip()
    
    # 4) Final fallback (same as before)
    logging.info(f"No equation resolved for property '{prop_name}', model '{model.get('name')}'")
    return ""

# -------------------------
# Core: build XML per component (UPDATED)
# -------------------------

def build_xml(json_data, uom_map, tempdep_config, model_eq_map, filename):
    comp = ET.Element("comp")

    # Extract TRCID (refCode) from filename
    trcid = REF_CODE_DEFAULT
    if filename:
        candidate = filename.split('_')[0]
        if candidate.isdigit():
            trcid = candidate

    # Basic identifiers
    aliases = json_data.get("aliases", [])
    name = json_data.get("name", "") or ""
    id_text = " ".join([a.upper() for a in aliases if a]).strip()
    if not id_text:
        id_text = "UNKNOWN"
    sim_id = str(json_data.get("SIMSCIID", "") or "")
    raw_cas = json_data.get("CASNO", "") or json_data.get("CAS", "")
    formatted_cas = normalize_cas(raw_cas)
    # if formatted_cas and not validate_cas_checksum(formatted_cas): logging.warning(f"{filename}: INVALID CAS {formatted_cas}"); formatted_cas=""
    if formatted_cas and not validate_cas_checksum(formatted_cas):

        INVALID_CAS.append({

            "File": filename,

            "Component": name,

            "SIMSCIID": sim_id,

            "CASNO": formatted_cas,

            "TRCID": trcid

        })

        logging.warning(

            f"{filename}: INVALID CAS {formatted_cas}"

        )




    if formatted_cas is None:
        logging.warning(f"{filename}: CAS '{raw_cas}' could not be normalized.")
        formatted_cas = str(raw_cas)

    formula = json_data.get("properties", {}).get("MOLFORMULA", "") or json_data.get("properties", {}).get("Formula", "")

    id_attrs = {
        "number": sim_id, "name": name, "formula": formula if formula else "",
        "casNum": formatted_cas if formatted_cas else "", "refCode": trcid,
    }
    id_el = ET.SubElement(comp, "id", id_attrs)
    id_el.text = id_text

    # Flags and structural properties
    properties = json_data.get("properties", {}).copy()
    applicable_phases_str = properties.pop("Applicable Phases", None) or properties.pop("ApplicablePhases", None)
    phase_flag_str = map_phases_to_flag(applicable_phases_str) if applicable_phases_str else "VLS"
    ET.SubElement(comp, "flags", {"phase": phase_flag_str})

    smiles_val = properties.pop("SMILES", "") or properties.pop("smiles", "")
    smiles_el = ET.SubElement(comp, "smiles")
    smiles_el.text = smiles_val if smiles_val else ""

    structure_val = json_data.get("structure", "")
    structure_el = ET.SubElement(comp, "structure")
    structure_el.text = structure_val if structure_val else ""

    # Fixed properties
    skip_keys = {"temp-dep_properties", "temp-dep_properties".lower(), "aliases", "smiles", 
                "MOLFORMULA", "Applicable Phases", "ApplicablePhases"}
    for k, v in properties.items():
        if v is None or k in skip_keys:
            continue
        attr_name = json_key_to_xml_attr_name(k)
        if isinstance(v, list):
            val = ", ".join(str(i) for i in v)
        else:
            val = str(v)
        prop_attrs = {"name": attr_name, "value": val}
        uom = uom_map.get(attr_name)
        if uom and str(uom).lower() not in ("", "dimensionless", "nan"):
            prop_attrs["uom"] = uom
        ET.SubElement(comp, "property", prop_attrs)

    # Temperature-dependent properties (UPDATED - Robust equation + Always print params)
    temp_props = json_data.get("temp-dep_properties", {}) or json_data.get("temp-dep_properties".lower(), {})
    if not temp_props:
        for alt_key in ("temp-dep-properties", "temp_dep_properties", "tempdep_properties"):
            if alt_key in json_data:
                temp_props = json_data.get(alt_key) or {}
                break

    for prop_name, models in temp_props.items():
        # Find matching config
        prop_key = next((k for k in tempdep_config.keys() if k.strip().lower() == prop_name.strip().lower()), None)
        cfg = tempdep_config.get(prop_key, {})

        tempUnit = cfg.get("tempUnit", "")
        propUnit = cfg.get("propUnit", "")
        moleWtBasis = cfg.get("moleWtBasis", "")
        logBasis = cfg.get("logBasis", "")

        # Process each model
        for model in models:
            tmin = model.get("tmin")
            tmax = model.get("tmax")
            corr_name = model.get("name", "")

            # NEW: Single-line robust equation resolution
            equation_val = resolve_equation(model, prop_name, model.get('equation'), model_eq_map, tempdep_config)
            print(f"Debug: Processing property '{prop_name}', model '{corr_name}', equation '{equation_val}'")

            # Prepare correlation attributes
            corr_attrib = {
                "name": prop_name, "equation": str(equation_val),
                "tempUnit": tempUnit if tempUnit else "",
                "propUnit": propUnit if propUnit else "",
                "tMin": str(tmin) if tmin is not None else "",
                "tMax": str(tmax) if tmax is not None else "",
            }

            # Add optional attributes
            def safe_xml_attr(val):
                if val is None:
                    return None
                s = str(val).strip().lower()
                if s in ("", "nan"):
                    return None
                return str(val).strip()

            mole_basis_val = safe_xml_attr(moleWtBasis)
            log_basis_val = safe_xml_attr(logBasis)
            if mole_basis_val:
                corr_attrib["moleWtBasis"] = mole_basis_val
            if log_basis_val:
                corr_attrib["logBasis"] = log_basis_val

            corr_elem = ET.SubElement(comp, "correlation", corr_attrib)
            lines = []

            # Parameter handling (ALL existing transformations preserved)
            # constants = model.get("constants", []) or []
            # parameters = model.get("parameters", []) or []
            constants = model.get("constants", []) or []

            raw_params = model.get("parameters", [])

            if raw_params is None:
                parameters = []
            elif isinstance(raw_params, list):
                parameters = raw_params
            else:
                parameters = [raw_params]
            print("LiquidDensity runtime parameters:", model.get("parameters"))
            if parameters and len(parameters) > 0:
                try:
                    # params_copy = [p for p in parameters]
                    params_copy = parameters if isinstance(parameters, list) else [parameters] if parameters is not None else []

                    # ALL your existing transformations (unchanged)
                    if corr_attrib.get("name") == "SurfaceTension" and equation_val in ("30"):
                        if len(params_copy) >= 1:
                            params_copy[0] = math.exp(params_copy[0])
                        else:
                            params_copy = []

                    if corr_attrib.get("name") == "LiquidViscosity" and equation_val in ("20"):
                        if len(params_copy)>0:
                            new_params = [0.0] * 10
                            p0 = params_copy[0] if len(params_copy) > 0 else 0.0
                            p1 = params_copy[1] if len(params_copy) > 1 else 0.0
                            p2 = params_copy[2] if len(params_copy) > 2 else 0.0
                            p3 = params_copy[3] if len(params_copy) > 3 else 0.0
                            new_params[0] = p0; new_params[1] = p1; new_params[2] = 0.0
                            new_params[3] = p3; new_params[4] = -3.0; new_params[5] = 0.0
                            new_params[6] = 0.0; new_params[7] = p2; new_params[8] = 0.0
                            new_params[9] = 0.0
                            params_copy = new_params
                        else:
                            params_copy = []

                    if corr_attrib.get("name") == "LiquidViscosity" and equation_val in ("13"):
                        if len(params_copy) > 1:
                            reordered = params_copy[1:] + params_copy[:1]
                            params_copy = reordered

                    if corr_attrib.get("name") == "VaporPressure" and equation_val in ("48"):
                        if len(params_copy) > 1:
                            params_copy = params_copy[1:]
                        else:
                            params_copy = []

                    # if corr_attrib.get("name") == "LiquidDensity" and str(equation_val) in ("61"):
                    #     if len(params_copy) > 1:
                    #         params_copy = params_copy[1:]
                    #     else:
                    #         params_copy = []

                    if corr_attrib.get("name") == "LiquidDensity":
                                            # Only apply the parameter slice for Equation 61
                        if str(equation_val) == "61":
                            if len(params_copy) > 1:
                                    params_copy = params_copy[1:]
                            else:
                                params_copy = []
                            # For Equation 1 (your JSON) and others, we keep params_copy as is
                        else:
                            pass

                    if corr_attrib.get("name") in ("IdealEnthalpy", "LiquidEnthalpy", "SolidEnthalpy"):
                        if corr_name in ("PolynomialHC", "Yaws.PolynomialExpansion", "DIPPR.PolynomialExpansion") and equation_val == "1":
                            if len(params_copy) > 0:
                                new_params = [0.0]
                                new_params += [params_copy[i] / (i + 1) for i in range(len(params_copy))]
                                params_copy = new_params
                            else:
                                params_copy = []
                        else:
                            if len(params_copy) > 0:
                                new_params = [0.0]
                                new_params += [params_copy[i] for i in range(len(params_copy))]
                                params_copy = new_params
                            else:
                                params_copy = []

                    if corr_attrib.get("name") in ("IdealEnthalpy", "IdealGasCp") and equation_val in ("40"):
                        try:
                            if len(params_copy) > 2:
                                params_copy[2] = float(params_copy[2]) * 1.0e5
                            if len(params_copy) > 6:
                                params_copy[6] = float(params_copy[6]) * 1.0e6
                        except Exception:
                            pass
                    
                    # CRITICAL: ALWAYS PRINT PARAMETERS IF THEY EXIST
                    param_line = " ".join(str(p) for p in params_copy)
                    if param_line.strip():
                        lines.append(param_line)
                        
                except Exception as e:
                    logging.error(f"Parameter processing failed for {prop_name}: {e}")

            corr_elem.text = "\n ".join(lines) + ("\n" if lines else "")

            comp_identifier = f"{name} ({formatted_cas})" if formatted_cas else name or filename
            logging.info(f"[{comp_identifier}] property '{prop_name}' model '{corr_name}' → equation '{equation_val}'")

    return comp

# -------------------------
# Convert function (UPDATED)
# -------------------------
def convert_all_json_to_xml(PROCESSEDDIR):
    uom_map = load_uom_mapping(CONFIG_FILE)
    tempdep_config = load_tempdep_property_config(CONFIG_FILE)
    model_eq_map = load_model_equation_mapping(CONFIG_FILE)  # NEW

    for filename in os.listdir(str(PROCESSEDDIR)):
        if not filename.lower().endswith(".json"):
            continue
       
        json_path = (PROCESSEDDIR/ filename)
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                json_data = json.load(f)
        except Exception as e:
            logging.error(f"Failed to load JSON {json_path}: {e}")
            continue

        xml_root = build_xml(json_data, uom_map, tempdep_config, model_eq_map, filename)  # Pass model_eq_map
        if xml_root is None:
            continue
        xml_str = pretty_xml(xml_root)

        # Safe filename
        # raw_name = (json_data.get("aliases", [None])[0] or json_data.get("name", "Component"))
        # raw_name = raw_name.strip(" ,-").upper().replace(" ", "_")
        # safe_base = re.sub(r'[<>:"/\\|?*]', '_', raw_name)
        # safe_base = safe_base[:80]
        # xml_filename = f"Comp-NIST-{safe_base}.xml"
        # Use TRCID directly for naming
        # Extract TRCID from filename
        trcid = filename.split('_')[0] if filename.split('_')[0].isdigit() else "UNKNOWN"
        safe_base = f"NST{trcid}"
        xml_filename = f"Comp-NIST-{safe_base}.xml"
        xml_path = (XML_OUTPUT_DIR/ xml_filename)
        
        try:
            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(xml_str)
            print(f"Converted {filename} → {xml_filename}")
        except Exception as e:
            logging.error(f"Failed to write XML {xml_path}: {e}")

# -------------------------
# Main execution
# -------------------------
if __name__ == "__main__":
    if MASTER_DIR.exists():
        convert_all_json_to_xml(MASTER_DIR)
    else:
        logging.warning(f"Processed dir not found: {MASTER_DIR}")

    if ASSIGNED_DIR.exists():
        convert_all_json_to_xml(ASSIGNED_DIR)
    else:
        logging.warning(f"Processed dir not found: {ASSIGNED_DIR}")
    
    if INVALID_CAS:

        invalid_df = pd.DataFrame(
            INVALID_CAS
        )

        invalid_file = (
            XML_OUTPUT_DIR
            / "Invalid_CAS.csv"
        )

        invalid_df.to_csv(

            invalid_file,

            index=False

        )

        print(

            f"\nInvalid CAS saved: "

            f"{invalid_file}"

        )

    logging.info("XML generation completed.")

"""
Script:
    11_transform_nist_json_to_simsci_format.py

Purpose:
    Transform NIST JSON component data into SimSci-compatible
    JSON format with mapped properties and SIMSCI IDs.

Functionality:
    - Reads NIST JSON files
    - Maps fixed and temperature-dependent properties
    - Assigns aliases, SMILES, and SIMSCI IDs
    - Generates processed JSON files
    - Separates matched and unmatched components
    - Creates logs for unmatched and duplicate components

Input:
    - NIST JSON files
    - Property mapping templates
    - Master component list
    - Library extraction data
    - SMILES data

Output:
    - Processed SimSci JSON files
    - Unmatched component reports
    - Log files
"""

import os
import json
import pandas as pd
import shutil
import logging
import re
import ast


# ==== CONFIG ====

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# PREREQUISITE DIRECTORIES
# ==================================================
PREREQ_DIR = BASE_DIR / "prerequisites"

EXCEL_INPUT_DIR = PREREQ_DIR / "excel_inputs"

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR

JSON_DIR = OUTPUT_DIR / "json"

SMILES_DIR = OUTPUT_DIR / "smiles"

PROCESSED_DIR = ( OUTPUT_DIR / "processed" / "full_library" / "1_components_Inmaster_withsimsciid")

LOG_DIR = BASE_DIR / "logs"

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================
MASTER_FILE = ( EXCEL_INPUT_DIR / "5_Master_Component_List.xlsx")

LIB_XML_EXTRACT_FILE = ( OUTPUT_DIR / "processed" / "full_library" / "2_Libraries_XML_Component_Extract.xlsx")

CONFIG_FILE = (EXCEL_INPUT_DIR / "6_NIST_Property_Mappings_Template.xlsx")

SMILES_FILE = (SMILES_DIR / f"2_compounds_smiles_{RUN_YEAR}_removedblanks.xlsx")

UNMATCHED_FILE = ( JSON_DIR / "unmatched_components.xlsx"
)

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

# logging.basicConfig(filename="property_extraction.log", level=logging.INFO)
logging.basicConfig(
    filename=LOG_DIR / "temp_dep_extraction.log",
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"

)
debug_logger = logging.getLogger("VCDEBUG")
debug_logger.setLevel(logging.DEBUG)

fh = logging.FileHandler(LOG_DIR / "vc_debug.log",mode="w")
fh.setLevel(logging.DEBUG)
fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
fh.setFormatter(fmt)
debug_logger.addHandler(fh)


# ==== HELPER FUNCTIONS ====

def normalize_corr_phase(constraints):

    if not constraints:
        return []

    normalized = set()

    for ph in constraints:
        ph = str(ph).strip().upper()

        if ph in ["V", "L", "G"]:
            normalized.add("VL")

        elif ph in ["S", "C1"]:
            normalized.add("S")

        elif ph == "VS":
            normalized.add("VLS")

        elif ph in ["VL", "LS", "VLS"]:
            normalized.add(ph)

    return list(normalized)

def normalize_cas(cas):
    """Normalize CAS to a digit-only string without leading zeros."""
    if pd.isna(cas):
        return None
    digits = ''.join(ch for ch in str(cas).strip() if ch.isdigit())
    return digits.lstrip('0') or '0'

def find_single_value_fixed_property(data, code_key):
    """Recursively search for a fixed-value property by its NIST code."""
    if isinstance(data, dict):
        if code_key in data and isinstance(data[code_key], (int, float, str, list)):
            return data[code_key]
        if "prop" in data and isinstance(data["prop"], dict) and data["prop"].get("code") == code_key:
            for model in data.get("models", []):
                if str(model.get("name", "")).strip().lower() == "single value":
                    params = model.get("parameters", [])
                    if isinstance(params, list):
                        if len(params) == 1 and isinstance(params[0], (int, float, str)):
                            return params[0]
                        if len(params) == 1 and isinstance(params[0], list) and len(params[0]) == 1:
                            inner = params[0]
                            if isinstance(inner, (int, float, str)):
                                return inner
                    elif isinstance(params, (int, float, str)):
                        return params
                # Always try PolynomialPressure as fallback
                if str(model.get("name", "")).strip().lower() == "polynomialpressure":
                    constants = model.get("constants", [])
                    if isinstance(constants, list) and len(constants) > 1:
                        return constants[1]
        for v in data.values():
            found = find_single_value_fixed_property(v, code_key)
            if found is not None:
                return found
    elif isinstance(data, list):
        for item in data:
            found = find_single_value_fixed_property(item, code_key)
            if found is not None:
                return found
    return None

def find_constraints(data, code_key):
    """Recursively find constraints list for a given property code."""
    if isinstance(data, dict):
        if "prop" in data and isinstance(data["prop"], dict) and data["prop"].get("code") == code_key:
            constraints = data.get("constraints", None)
            if isinstance(constraints, list):
                return constraints
        for v in data.values():
            found = find_constraints(v, code_key)
            if found is not None:
                return found
    elif isinstance(data, list):
        for item in data:
            found = find_constraints(item, code_key)
            if found is not None:
                return found
    return None

def gather_all_constraints(data):
    """Aggregate all unique constraints from 'datasets' list in JSON."""
    datasets = data.get("datasets", [])
    constraints = set(str(c) for ds in datasets for c in ds.get("constraints", []))
    return sorted(constraints) if constraints else None

def get_temp_max_min(ranges):
    tmin, tmax = None, None
    if not ranges or not isinstance(ranges, list):
        return tmin, tmax
    for r in ranges:
        if r.get("code", "").strip().upper() == "T":
            tmin = r.get("min")
            tmax = r.get("max")
            break
    return tmin, tmax

def normalize_phase(phase_str):
    if not phase_str or pd.isna(phase_str):
        return None
    phase_str = str(phase_str).strip().upper()
    mapping = {
        "LIQUID": "L", "L": "L",
        "SOLID": "C1", "C1": "C1",
        "GAS": "G", "G": "G",
        "IDEAL GAS": "IG", "IG": "IG",
        "VAPOR": "G", "G": "G"
    }
    return mapping.get(phase_str, phase_str)


def normalize_model_name(name: str) -> str:
    """Normalize model names for robust comparison."""
    return re.sub(r"[\s.]+", "", str(name).strip().lower())


def extract_temperature_dependent_properties(json_data, tdep_df):

    # ---------------- Get Component Name ----------------
    comp_name = ""

    # 1) Top level name
    if json_data.get("name"):
        comp_name = json_data["name"].strip()

    # 2) Inside compounds list
    elif "compounds" in json_data and isinstance(json_data["compounds"], list):
        comp_block = json_data["compounds"][0]
        if isinstance(comp_block, dict):

            # 2a) names array
            if "names" in comp_block and isinstance(comp_block["names"], list) and comp_block["names"]:
                comp_name = comp_block["names"][0].strip()

            # 2b) single name field
            elif comp_block.get("name"):
                comp_name = comp_block["name"].strip()

    # 3) Fallback to formula
    if not comp_name and json_data.get("formula"):
        comp_name = json_data["formula"]

    # 4) Default
    if not comp_name:
        comp_name = "UNKNOWN"

    comp_id = None

    # TRCID inside "compounds"
    if "compounds" in json_data and isinstance(json_data["compounds"], list) and json_data["compounds"]:
        comp_id = json_data["compounds"][0].get("trcid")

    # Other TRCID formats
    if not comp_id:
        trc = json_data.get("trcid")
        if isinstance(trc, list) and trc:
            comp_id = trc[0]
        elif isinstance(trc, (int, str)):
            comp_id = trc

    if not comp_id:
        comp_id = "UNKNOWN"

    extracted_props = {}

    def normalize_model_name(name):
        return re.sub(r"[\s.]+", "", str(name).strip().lower()) 


    # ----------------- Loop Config Rows -----------------
    for _, row in tdep_df.iterrows():
        simsci_name = str(row.get("SimSci_Name", "")).strip()

        nist_tokens = [
            t.strip().upper() for t in str(row.get("NIST_Name", "")).split("/") if t.strip()
        ]

        # model_name_cells = [m.strip() for m in str(row.get("Model_Name", "")).split("/")]
        # equation_cells = [e.strip() for e in str(row.get("Equation", "")).split("/")]
        # equation_cells = [e for e in equation_cells if e] # remove empty strings
        # if not equation_cells:
        #     continue

        # ================= SAFE MODEL / EQUATION PARSING =================

        import re

        def split_models(model_str):
            # Split on "/" but ignore slash inside single quotes
            parts = re.split(r"/(?=(?:[^']*'[^']*')*[^']*$)", str(model_str))
            return [p.strip().strip("'") for p in parts if p.strip()]


        def split_equations(eq_str):
            return [e.strip() for e in str(eq_str).split("/") if e.strip()]


        model_name_cells = split_models(row.get("Model_Name", ""))
        equation_cells = split_equations(row.get("Equation", ""))

        # ---- Strict validation ----
        if not model_name_cells:
            continue

        if not equation_cells:
            continue

        if len(model_name_cells) != len(equation_cells):
            print(f"\n[ERROR] Model–Equation mismatch for property: {sim_name}")
            print(f"Models    ({len(model_name_cells)}): {model_name_cells}")
            print(f"Equations ({len(equation_cells)}): {equation_cells}")
            continue  # skip safely

        # Create safe mapping
        model_equation_map = dict(zip(model_name_cells, equation_cells))

        phase_raw = str(row.get("Phase","")).strip()
        try:
            phases = ast.literal_eval(phase_raw) if phase_raw.startswith("[") else [phase_raw]
        except:
            phases = [phase_raw]
        phases = [p.strip().upper() for p in phases if p.strip()]

        chosen_model = None
        chosen_constraints = None
        chosen_eq = None
        chosen_nist_token = None

        # -------- Get all JSON models for NIST token --------
        def get_candidates(n_token):
            candidate_list = []
            for ds in json_data.get("datasets", []):
                if str(ds.get("prop", {}).get("code", "")).upper() != n_token:
                    continue

                var_codes = [v.get("code","").upper() for v in ds.get("variables",[])]
                if "T" not in var_codes:
                    continue

                ds_const = [str(c).upper() for c in ds.get("constraints",[]) if str(c).strip()]

                for m in ds.get("models",[]):
                    m_const = [str(c).upper() for c in m.get("constraints",[]) if str(c).strip()]

                    # preserve JSON phase order, avoid set()
                    all_const = []
                    for c in ds_const + m_const:
                        if c not in all_const:
                            all_const.append(c)

                    m_norm = normalize_model_name(m.get("name",""))

                    # Skip cubic everywhere
                    if "cubic" in m_norm or "spline" in m_norm:
                        logging.info(f"[{comp_name}|{comp_id}] SKIP Cubic {m.get('name')} for {simsci_name}")
                        continue

                    candidate_list.append((ds, m, all_const))

            return candidate_list

        # ------------------ Main Matching Loop ------------------
        for token in nist_tokens:

            all_candidates = get_candidates(token)
            if not all_candidates:
                continue

            # # Single phase priority (C1 / G / IG)
            # if len(phases) == 1:
            #     p = phases[0]
            #     exact = [(d,m,c) for d,m,c in all_candidates if set(c)=={p}]
            #     contains = [(d,m,c) for d,m,c in all_candidates if p in c]
            #     candidates = exact if exact else contains
            # else:
            #     exact = [(d,m,c) for d,m,c in all_candidates if c == phases]
            #     swapped = [(d,m,c) for d,m,c in all_candidates if c == phases[::-1]]
            #     overlapped = [(d,m,c) for d,m,c in all_candidates if set(phases)&set(c)]
            #     candidates = exact or swapped or overlapped
            
            # ================= STRICT PHASE MATCHING -2=================

            # if len(phases) == 1:
            #     # STRICT single-phase match only
            #     p = phases[0]
            #     candidates = [
            #         (d, m, c) for d, m, c in all_candidates
            #         if set(c) == {p}
            #     ]

            # else:
            #     # STRICT multi-phase match only (order insensitive)
            #     candidates = [
            #         (d, m, c) for d, m, c in all_candidates
            #         if sorted(c) == sorted(phases)
                # ]
            
            # ================= THERMODYNAMICALLY CORRECT PHASE MATCHING-3 =================

            if len(phases) == 1:
                # Single-phase property (L, G, IG, C1)
                p = phases[0]
                candidates = [
                    (d, m, c) for d, m, c in all_candidates
                    if p in c
                ]

            else:
                # Two-phase property (e.g., ["L","G"])
                # Require ALL template phases to be present
                candidates = [
                    (d, m, c) for d, m, c in all_candidates
                    if all(p in c for p in phases)
                ]

            if not candidates:
                continue

            # ------ Try Excel Model_Name match ------
            # match_found = False
            # for ds, m, c in candidates:
            #     m_norm = normalize_model_name(m.get("name",""))
            #     match_idx = None

            #     for j, cfg_m in enumerate(model_name_cells):
            #         if not cfg_m: continue
            #         cfg_norm = normalize_model_name(cfg_m)
            #         if m_norm == cfg_norm or m_norm in cfg_norm or cfg_norm in m_norm:
            #             match_idx = j; break

            #     if match_idx is not None and match_idx < len(equation_cells):
            #         chosen_model = m
            #         chosen_constraints = c
            #         chosen_eq = equation_cells[match_idx]
            #         chosen_nist_token = token
            #         match_found = True
            #         break


            match_found = False
            for ds, m, c in candidates:

                m_norm = normalize_model_name(m.get("name", ""))
                matched_model_name = None

                for cfg_m in model_name_cells:
                    if not cfg_m:
                        continue

                    cfg_norm = normalize_model_name(cfg_m)

                    if m_norm == cfg_norm or m_norm in cfg_norm or cfg_norm in m_norm:
                        matched_model_name = cfg_m
                        break

                # ---- SAFE EQUATION FETCH ----
                if matched_model_name:
                    chosen_eq = model_equation_map.get(matched_model_name)

                    if not chosen_eq:
                        print(f"[WARNING] Equation missing for model {matched_model_name} in {sim_name}")
                        continue  # skip safely

                    chosen_model = m
                    chosen_constraints = c
                    chosen_nist_token = token
                    match_found = True
                    break

            if match_found:
                break

            # ------ Fallback (skip cubic already done earlier) ------
            for ds, m, c in candidates:
                m_norm = normalize_model_name(m.get("name",""))
                chosen_model = m
                chosen_constraints = c
                chosen_nist_token = token

                match_idx = None
                for j, cfg_m in enumerate(model_name_cells):
                    if not cfg_m: continue
                    cfg_norm = normalize_model_name(cfg_m)
                    if m_norm == cfg_norm or m_norm in cfg_norm or cfg_norm in m_norm:
                        match_idx = j; break

                # chosen_eq = equation_cells[match_idx] if match_idx is not None else equation_cells[0]
                if match_idx is not None and 0 <= match_idx < len(equation_cells):
                    chosen_eq = equation_cells[match_idx]
                else:
                    chosen_eq = equation_cells
            else:
                chosen_eq = None # or "" if you prefer a string

                logging.info(f"[{comp_name}|{comp_id}] Fallback -> {m.get('name')} for {simsci_name}")
                match_found = True
                break

            if match_found:
                break

        # ---------------- Store Result ----------------
        # if chosen_model:
        if chosen_model and chosen_eq and str(chosen_eq).strip():
            ranges = chosen_model.get("ranges",[])
            tmin = ranges[0].get("min") if ranges else None
            tmax = ranges[0].get("max") if ranges else None

            extracted_props.setdefault(simsci_name, []).append({
                "name": chosen_model.get("name"),
                "constants": chosen_model.get("constants", []),
                "parameters": chosen_model.get("parameters", []),
                "tmin": tmin, "tmax": tmax,
                "constraints": sorted(normalize_corr_phase(chosen_constraints)),
                "equation": chosen_eq
            })

            logging.info(
                f"[{comp_name}|{comp_id}] Selected {chosen_model.get('name')} for {simsci_name} "
                f"(Eq={chosen_eq}, ConfigPh={phases}, JSONPh={normalize_corr_phase(chosen_constraints)}, NIST={chosen_nist_token})"
            )
        else:
            logging.warning(
                f"[{comp_name}|{comp_id}] No model match for {simsci_name} NIST={nist_tokens} ConfigPh={phases}"
            )

    return extracted_props, comp_id


def get_vc_from_tdep(temp_dep_props, mmass):
    if not temp_dep_props or not mmass:
        debug_logger.debug(f"No temp_dep_props or mmass. mmass={mmass}")
        return None

    # Identify Liquid Density keys
    models = None
    for key in ["LiquidDensity", "VDN"]:
        if key in temp_dep_props:
            models = temp_dep_props[key]
            break

    if not models:
        debug_logger.debug("No VDN/LiquidDensity models found")
        return None

    # Phase priority: LG > GL > L alone
    phase_priority = [["L","G"], ["G","L"], ["L"]]

    def phase_rank(c):
        if c in phase_priority:
            return phase_priority.index(c)
        if isinstance(c, list) and set(c) in [set(p) for p in phase_priority]:
            return [set(p) for p in phase_priority].index(set(c))
        return 99

    # Sort by phase priority only
    sorted_models = sorted(models, key=lambda m: phase_rank(m.get("constraints",[])))

    for m in sorted_models:
        name = str(m.get("name","")).replace(" ", "").lower()
        eq = str(m.get("equation",""))
        params = m.get("parameters", [])
        cons = m.get("constraints", [])

        debug_logger.debug(f"Model check name={name}, eq={eq}, cons={cons}, params={params}")

        if eq != "61":
            continue
        if not ("vdn" in name or "vdns" in name or "vdnsexp" in name):
            continue
        # if "L" not in cons:  # Must have liquid
        #     continue
            # Relaxed phase check (FIX)
    # ---- SAFE PHASE FIX ----
        if any(p in cons for p in ["L", "G"]):
            pass
        else:
            has_liquid_model = any(
                any(p in mm.get("constraints", []) for p in ["L", "G"])
                for mm in sorted_models
            )
            if has_liquid_model:
                continue

        # YOUR LOGIC: Always use the FIRST parameter
        try:
            first = params[0]
            vc = float(mmass) / float(first)
            debug_logger.info(f"VC computed = {vc} using first param={first}, MW={mmass}")
            return vc
        except Exception as e:
            debug_logger.error(f"VC calc failed: {e}, params={params}")

    debug_logger.warning("No valid VDN model for VC")
    return None


# ==== MAIN PROCESSING ====

if __name__ == "__main__":
    # === Load config Excel sheets ===
    master_df = pd.read_excel(MASTER_FILE, header=5)
    cas_col = next((c for c in master_df.columns if 'CAS' in c.upper()), None)
    name_col = (next((c for c in master_df.columns if c.strip().lower() == "name"), None)
                or next((c for c in master_df.columns if "NAME" in c.upper() and "UNNAMED" not in c.upper()), None)
                or next((c for c in master_df.columns if "NAME" in c.upper()), None))
    id_col = next((c for c in master_df.columns if "SIMSCI" in c.upper()), None)
    alias_cols = [c for c in master_df.columns if "NAME" in c.upper() and c != name_col]
    # --- Load secondary XML extract source ---
    lib_xml_df = pd.read_excel(LIB_XML_EXTRACT_FILE)
    lib_xml_df.columns = lib_xml_df.columns.str.strip()

    # Normalize CAS
    master_df["_CAS_NORM"] = master_df[cas_col].apply(normalize_cas)
    lib_xml_df["_CAS_NORM"] = lib_xml_df["CASNO"].apply(normalize_cas)

    # CAS lookup sets
    master_cas_set = set(master_df["_CAS_NORM"].dropna())
    lib_cas_set = set(lib_xml_df["_CAS_NORM"].dropna())

    # Combined CAS knowledge (for matching only)
    known_cas_set = master_cas_set.union(lib_cas_set)

    fixed_df = pd.read_excel(CONFIG_FILE, sheet_name="FIXED_PROP_NAMES")
    tdep_df = pd.read_excel(CONFIG_FILE, sheet_name="TDEP_COR_NAMES")
    tdep_df["NIST_Name"] = tdep_df["NIST_Name"].astype(str).str.strip()
    tdep_df["SimSci_Name"] = tdep_df["SimSci_Name"].astype(str).str.strip()
    tdep_df["Model_Name"] = tdep_df["Model_Name"].astype(str).str.strip().str.lower()
    tdep_df["Phase"] = tdep_df["Phase"].apply(normalize_phase)
    smiles_lookup = {}
    if os.path.exists(SMILES_FILE):
        smiles_df = pd.read_excel(SMILES_FILE)
        smiles_df = smiles_df.dropna(subset=['CASRN', 'SMILES'])
        smiles_df = smiles_df[(smiles_df['CASRN'].astype(str).str.strip() != '')]
        smiles_lookup = {normalize_cas(row['CASRN']): str(row['SMILES']).strip() for _, row in smiles_df.iterrows()}
    unmatched_list = []
    skipped_list = []   
    duplicates_list = []
    default_id_counter = 2000000
    seen_cas = set()
    for idx, row in master_df.iterrows():
        cas_raw = row.get(cas_col, None)
        cas_str = "" if pd.isna(cas_raw) else str(cas_raw).strip()
        if not cas_str:
            continue
        # print("Hi cas_str")
        comp_name = row.get(name_col, "")
        comp_name = "" if pd.isna(comp_name) else str(comp_name).strip()
        sim_id = row.get(id_col, None)
        if sim_id is None:
            sim_id = default_id_counter
            default_id_counter += 1
        cas_digits = normalize_cas(cas_str)
        # print(f"Processing CAS: {cas_str} (Digits: {cas_digits})")
        if cas_digits in seen_cas:
            duplicates_list.append({"CAS": cas_str, "Name": comp_name, "SimSci_ID": sim_id})
            continue
        else:
            seen_cas.add(cas_digits)
        # Extract aliases, remove duplicates, skip blank/nan, skip main name
        # aliases = []
        # seen_alias = set()
        # for c in alias_cols:
        #     val = row.get(c, None)
        #     if val is None or pd.isna(val):
        #         continue
        #     sval = str(val).strip()
        #     if sval and sval.lower() != "nan" and sval != comp_name and sval not in seen_alias:
        #         aliases.append(sval)
        #         seen_alias.add(sval)


        # Find corresponding JSON file
        json_file = None
        if os.path.isdir(JSON_DIR):
            for fname in os.listdir(JSON_DIR):               
                parts = fname.split('_CASNO_')
                if len(parts) == 2 and parts[1].endswith('.json'):
                    file_casrn = parts[1].replace('.json', '')
                    if cas_digits == file_casrn:
                        json_file = os.path.join(JSON_DIR, fname)
                        print(f"Found matching JSON file: {json_file}")
                        break
        if not json_file:
            unmatched_list.append({"CAS": cas_str, "Name": comp_name, "SimSci_ID": sim_id})
            continue        
        if cas_digits not in known_cas_set:
    # CAS not found in Master or XML Extract → follow old not-in-master flow
            continue
        source_used = ("MASTER" if cas_digits in master_cas_set else "XML_EXTRACT")

        logging.info(f"[CAS={cas_digits}] Source used: {source_used}")

        with open(json_file, "r", encoding="utf-8") as jf:
            data = json.load(jf)
            # print(f"Loaded JSON data from {json_file}")       
         # === Temperature-dependent properties
        temp_dep_props, comp_id = extract_temperature_dependent_properties(data, tdep_df)
        
        mapped_props = {
            "CASNO": cas_str,
            "name": comp_name,
            "SIMSCIID": sim_id,
            "properties": {}
        }
        #Extract aliases  format NIST_TRCID
        aliases = []
        seen_alias = set()
        # NEW: Use TRCID as primary alias in {NIST_TRCID} format
        aliases = []
        if comp_id != "UNKNOWN":
            aliases.append(f"NST{comp_id}")
        else:
            seen_alias = set()
            for c in alias_cols:
                val = row.get(c, None)
                if val is None or pd.isna(val): continue
                sval = str(val).strip()
                if sval and sval.lower() != 'nan' and sval != comp_name and sval not in seen_alias:
                    aliases.append(sval)
                    seen_alias.add(sval)
        if aliases:
            mapped_props["aliases"] = aliases
            
        compounds = data.get("compounds", [])
        first_comp = compounds[0] if isinstance(compounds, list) and compounds and isinstance(compounds[0], dict) else {}

        mapped_props["temp-dep_properties"] = temp_dep_props
        # print( temp_dep_props)
        # === FIXED properties ===
        for _, cfg in fixed_df.iterrows():
            nist_name = str(cfg["NIST_Name"]).strip()
            sim_name = str(cfg["SimSci_Name"]).strip()
            # print(sim_name)
            conv = float(cfg["Conv_Factor"]) if not pd.isna(cfg["Conv_Factor"]) else 1.0
            raw_val = None
            # VC extracted, skip fixed lookup
            # print("Before VC if")
            # print(f'Repr of sim_name: {repr(sim_name)}')
            # print(f'sim_name.lower() == "vc": {sim_name.lower() == "vc"}')
            if sim_name.lower() == "vc":
                # print("enterd fisrt if")
                mmass = (
                    mapped_props["properties"].get("mmass")
                    or mapped_props["properties"].get("MW")
                    or first_comp.get("mmass")
                )
                # print(mmass)

                debug_logger.info(f"[VC CHECK] Component={comp_name}, CAS={cas_digits}, MW={mmass}")
                # print("HI")
                vc_val = get_vc_from_tdep(temp_dep_props, mmass)
                # print(f"Hello {vc_val}")
                debug_logger.info(f"[VC RESULT] VC={vc_val}")

                if vc_val is not None:
                    mapped_props["properties"][sim_name] = vc_val * conv
                    continue  # ONLY skip fixed lookup if VC derived successfully

            # No VC from temp-dep, fallback to fixed lookup below
            if nist_name.lower() == "mmass" and first_comp:
                raw_val = first_comp.get("mmass", None)
            elif nist_name.lower() == "formula" and first_comp:
                raw_val = first_comp.get("formula", None)
            # elif sim_name.lower() == "applicable phases":
            #     all_constraints = gather_all_constraints(data)
            #     mapped_props["properties"][sim_name] = ", ".join(all_constraints) if all_constraints else None
            #     continue
            ########
            elif sim_name.lower() == "applicable phases":

                raw_constraints = gather_all_constraints(data)

                if not raw_constraints:
                    mapped_props["properties"][sim_name] = None
                else:
                    normalized = normalize_corr_phase(raw_constraints)
                    mapped_props["properties"][sim_name] = ", ".join(sorted(normalized))

                continue
            else:
                raw_val = find_single_value_fixed_property(data, nist_name)
            constraints_val = find_constraints(data, nist_name)
            if raw_val is not None:
                if isinstance(raw_val, (int, float)):
                    mapped_props["properties"][sim_name] = raw_val * conv
                else:
                    mapped_props["properties"][sim_name] = raw_val
            else:
                mapped_props["properties"][sim_name] = None
            if sim_name.lower() == "applicable phases" and constraints_val:
                try:
                    mapped_props["properties"][sim_name] = ", ".join(map(str, constraints_val))
                except Exception:
                    mapped_props["properties"][sim_name] = constraints_val
        # === SMILES mapping
        smiles_val = smiles_lookup.get(cas_digits, None)
        if smiles_val:
            mapped_props["properties"]["SMILES"] = smiles_val


        # Write output JSON using original filename
        out_file = os.path.join(PROCESSED_DIR, os.path.basename(json_file))
        with open(out_file, "w", encoding="utf-8") as out:
            json.dump(mapped_props, out, indent=2)
    # Write unmatched list if any
    if unmatched_list:
        pd.DataFrame(unmatched_list).to_excel(UNMATCHED_FILE, index=False)
    # Write skipped components (missing CAS)
    skipped_file = os.path.join(JSON_DIR, "skipped_components.xlsx")
    if skipped_list:
        pd.DataFrame(skipped_list).to_excel(skipped_file, index=False)
    # Write duplicate components
    duplicates_file = os.path.join(JSON_DIR, "duplicate_components.xlsx")
    if duplicates_list:
        pd.DataFrame(duplicates_list).to_excel(duplicates_file, index=False)
    # Process not processed JSON files
    NOT_PROCESSED_DIR = (OUTPUT_DIR/ "processed"/ "full_library"/ "2_components_notInmaster_nosimsciid")

    NOT_PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    all_json_files = []
    for fname in os.listdir(JSON_DIR):
        parts = fname.split('_CASNO_')
        if len(parts) == 2 and parts[1].endswith('.json'):
            file_casrn = parts[1].replace('.json', '')
            norm_cas = file_casrn.lstrip('0') or '0'
            all_json_files.append((fname, norm_cas))
    processed_cas_set = seen_cas
    for fname, cas_digits in all_json_files:
        if cas_digits in processed_cas_set:
            continue  # Already processed in main loop
        json_path = os.path.join(JSON_DIR, fname)
        with open(json_path, "r", encoding="utf-8") as jf:
            data = json.load(jf)
        compounds = data.get("compounds", [])
        first_comp_name = ""
        if isinstance(compounds, list) and compounds and isinstance(compounds[0], dict):
            first_comp = compounds[0]
            if "names" in first_comp and isinstance(first_comp["names"], list) and first_comp["names"]:
                first_comp_name = first_comp["names"][0].strip()
            else:
                first_comp_name = first_comp.get("name", "").strip()
        mapped_props = {
            "CASNO": cas_digits,
            "name": first_comp_name if first_comp_name else "UNKNOWN",
            "SIMSCIID": None,
            "properties": {}
        }
        # === Temperature-dependent properties
        temp_dep_props,comp_id = extract_temperature_dependent_properties(data, tdep_df)
        aliases = []
        if comp_id != "UNKNOWN":
            aliases.append(f"NST{comp_id}")        
        if aliases:
            mapped_props['aliases'] = aliases
        # you still have this if you want to use it later
        compounds = data.get("compounds", [])
        first_comp = compounds[0] if compounds and isinstance(compounds[0], dict) else {}

        mapped_props["temp-dep_properties"] = temp_dep_props
        for _, cfg in fixed_df.iterrows():
            nist_name = str(cfg["NIST_Name"]).strip()
            sim_name = str(cfg["SimSci_Name"]).strip()
            conv = float(cfg["Conv_Factor"]) if not pd.isna(cfg["Conv_Factor"]) else 1.0
            raw_val = None
            if sim_name.lower() == "vc":
                # print("enterd fisrt if")
                mmass = (
                    mapped_props["properties"].get("mmass")
                    or mapped_props["properties"].get("MW")
                    or first_comp.get("mmass")
                )
                # print(mmass)

                debug_logger.info(f"[VC CHECK] Component={comp_name}, CAS={cas_digits}, MW={mmass}")
                # print("HI")
                vc_val = get_vc_from_tdep(temp_dep_props, mmass)
                # print(f"Hello {vc_val}")
                debug_logger.info(f"[VC RESULT] VC={vc_val}")

                if vc_val is not None:
                    mapped_props["properties"][sim_name] = vc_val * conv
                    continue  # ONLY skip fixed lookup if VC derived successfully
            if nist_name.lower() == "mmass" and first_comp:
                raw_val = first_comp.get("mmass", None)
            elif nist_name.lower() == "formula" and first_comp:
                raw_val = first_comp.get("formula", None)
            elif sim_name.lower() == "applicable phases":
                all_constraints = gather_all_constraints(data)
                mapped_props["properties"][sim_name] = ", ".join(all_constraints) if all_constraints else None
                continue
            else:
                raw_val = find_single_value_fixed_property(data, nist_name)
            constraints_val = find_constraints(data, nist_name)
            if raw_val is not None:
                if isinstance(raw_val, (int, float)):
                    mapped_props["properties"][sim_name] = raw_val * conv

                else:
                    mapped_props["properties"][sim_name] = raw_val
            else:
                mapped_props["properties"][sim_name] = None
            if sim_name.lower() == "applicable phases" and constraints_val:
                try:
                    mapped_props["properties"][sim_name] = ", ".join(map(str, constraints_val))
                except Exception:
                    mapped_props["properties"][sim_name] = constraints_val
        # Extract temperature dependent properties
        temp_dep_props,comp_id = extract_temperature_dependent_properties(data, tdep_df)
        mapped_props["temp-dep_properties"] = temp_dep_props
        # === SMILES mapping
        smiles_val = smiles_lookup.get(cas_digits, None)
        if smiles_val:
            mapped_props["properties"]["SMILES"] = smiles_val
        # Write out JSON to NOT_PROCESSED_DIR using original filename
        out_file = os.path.join(NOT_PROCESSED_DIR, fname)
        with open(out_file, "w", encoding="utf-8") as outf: 
            json.dump(mapped_props, outf, indent=2)

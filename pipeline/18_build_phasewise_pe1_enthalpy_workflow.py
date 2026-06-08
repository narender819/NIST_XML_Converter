"""
Script:
    18_build_phasewise_pe1_enthalpy_workflow.py

Purpose:
    This script processes PE1 legacy output files, extracts
    phase-specific thermodynamic results, and builds separate
    LCP, ICP, and SCP workflow sheets.

Functionality:
    - Reads core thermodynamic property data
    - Reads processed component JSON files
    - Identifies component availability in SIMSCI libraries
    - Extracts enthalpy correlation availability flags
    - Parses PE1 legacy output files
    - Routes phase-specific properties into:
        * LCP sheet
        * ICP sheet
        * SCP sheet
    - Calculates HDeparture values for ICP workflow
    - Generates a consolidated multi-sheet Excel workflow report
    - Auto-adjusts Excel column widths

Input:
    - Core thermodynamic property Excel file
    - SIMSCI library extraction Excel file
    - Processed component JSON files
    - PE1 output files

Output:
    Multi-sheet Excel report containing phasewise PE1 enthalpy workflow data
"""



import pandas as pd
import json
import re
from pathlib import Path
from openpyxl import load_workbook

from pathlib import Path

# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR

PROCESSED_DIR = (
    OUTPUT_DIR
    / "processed"
    / "full_library"
)

PROPEVAL_RUNS_DIR = (
    OUTPUT_DIR
    / "propeval_runs"
)

# ==================================================
# INPUT FILES
# ==================================================
core_props_excel = (
    PROCESSED_DIR
    / f"12_NIST_Component_KeyThermoProperties_{RUN_YEAR}_TCPCAF_UPDATED.xlsx"
)

enthalpy_workflow_excel = (
    PROCESSED_DIR
    / f"14_NIST_Splitsheets_PE1legacy_Hdepart.xlsx"
)

simsci_excel = (
    PROCESSED_DIR
    / "2_Libraries_XML_Component_Extract.xlsx"
)

json_folders = [

    PROCESSED_DIR
    / "1_components_Inmaster_withsimsciid",

    PROCESSED_DIR
    / "3_components_notInmaster_assignedsimsciid"
]

# ==================================================
# PROPEVAL OUTPUTS
# ==================================================
pe1_dir = (
    PROPEVAL_RUNS_DIR
    / "NIST"
)

# ---------------------------------------------------
# PE1 Routing Config (NEW)
# ---------------------------------------------------
PE1_ROUTING_MAP = {
    # -------- ICP --------
    "PTP@TTP (Kpa)": "ICP",
    "Pc_Calc (Kpa)": "ICP",

    # -------- LCP --------
    "LDEN@60F (kg/m3)": "LCP",
    "LDEN@25C (kg/m3)": "LCP",
}



def auto_adjust_column_widths(excel_path, max_width=40, padding=2):
    wb = load_workbook(excel_path)

    for ws in wb.worksheets:
        for column_cells in ws.columns:
            max_len = 0
            col_letter = column_cells[0].column_letter

            for cell in column_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            adjusted = min(max_len + padding, max_width)
            ws.column_dimensions[col_letter].width = adjusted

    wb.save(excel_path)

# ---------------------------------------------------
# Helpers
# ---------------------------------------------------
# def normalize_cas(cas):
#     if pd.isna(cas):
#         return None
#     s = str(cas).strip()
#     # remove hyphens and handle float-like values
#     s = s.replace("-", "")
#     try:
#         f = float(s)
#         if f.is_integer():
#             return str(int(f))
#     except:
#         pass
#     return s

def normalize_cas(cas):
    if pd.isna(cas):
        return None

    s = str(cas).strip()

    # remove hyphens
    s = s.replace("-", "")

    # remove decimal part if present
    if "." in s:
        s = s.split(".")[0]

    return str(s)



def extract_last_float(line):
    nums = re.findall(r"[-+]?[\d.]+(?:[eE][-+]?\d+)?", line)
    return float(nums[-1]) if nums else None

def make_safe_name(alias, max_len=30):
    """Convert alias NST1143 → NST1143 (safe for filename)"""
    if pd.isna(alias): return ""
    name = str(alias).strip().upper().replace(" ", "_")
    name = re.sub(r"[^A-Z0-9_]", "", name)
    name = re.sub(r"_+", "_", name)
    return name[:max_len]

# ---------------------------------------------------
# SIMSCI CAS lookup
# ---------------------------------------------------

def build_simsci_cas_set(path, sheet_names):
    try:
        all_sheets = pd.read_excel(path, sheet_name=None)
        cas_set = set()

        for sheet_name in sheet_names:
            if sheet_name not in all_sheets:
                print(f"Sheet not found (skipped): {sheet_name}")
                continue

            df = all_sheets[sheet_name]

            for col in df.columns:
                if "CAS" in str(col).upper():
                    norm = df[col].dropna().apply(normalize_cas)
                    cas_set.update([c for c in norm if c is not None])

        return cas_set

    except Exception as e:
        print("SIMSCI read error:", e)
        return set()

# def load_enthalpy_flags_from_json(folders):
#     cas_map = {}
#     for folder in folders:
#         for jf in Path(folder).glob("*.json"):
#             try:
#                 data = json.loads(jf.read_text(encoding="utf-8"))
#                 # cas = normalize_cas(data.get("CASNO"))
#                 cas = str(normalize_cas(data.get("CASNO")))
#                 tprops = data.get("temp-dep_properties", {})

#                 if cas is None:
#                     continue  # skip bad CAS

#                 cas_map[cas] = {
#                     "LCP_Exists": "Yes" if "LiquidEnthalpy" in tprops else "No",
#                     "ICP_Exists": "Yes" if "IdealEnthalpy" in tprops else "No",
#                     "SCP_Exists": "Yes" if "SolidEnthalpy" in tprops else "No",
#                     "HVAP_Exists": "Yes" if "LatentHeat" in tprops else "No",
#                 }
#             except Exception as e:
#                 print(f"JSON parse error in {jf}: {e}")
#     return cas_map

def load_enthalpy_flags_from_json(folders):
    cas_map = {}
    for folder in folders:
        for jf in Path(folder).glob("*.json"):
            try:
                data = json.loads(jf.read_text(encoding="utf-8"))
                cas = str(normalize_cas(data.get("CASNO")))
                tprops = data.get("temp-dep_properties", {})
                
                if cas is None:
                    continue  # skip bad CAS

                # Helper to check if property exists (handles dict or list-of-dicts)
                def prop_exists(prop_name, props):
                    if isinstance(props, dict):
                        return prop_name in props
                    elif isinstance(props, list):
                        return any(prop_name.lower() in p.get("name", "").lower() for p in props)
                    return False

                cas_map[cas] = {
                    "LCP_Exists": "Yes" if prop_exists("LiquidEnthalpy", tprops) else "No",
                    "ICP_Exists": "Yes" if prop_exists("IdealEnthalpy", tprops) else "No",
                    "SCP_Exists": "Yes" if prop_exists("SolidEnthalpy", tprops) else "No",
                    "HVAP_Exists": "Yes" if prop_exists("LatentHeat", tprops) else "No",
                }
            except Exception as e:
                print(f"JSON parse error in {jf}: {e}")
    return cas_map



# ---------------------------------------------------
# PE1 parser
# ---------------------------------------------------
def parse_pe1(pe1_file):
    results = {}
    lines = Path(pe1_file).read_text(encoding="cp1252", errors="ignore").splitlines()
    i = 0
    while i < len(lines):
        l = lines[i]
        #Vapor
        if "HIG@NBP" in l:
            results["HIG@NBP (J/kg-mole)"] = extract_last_float(l) 
        elif "HV@NBP" in l or "HV@NBP(J" in l:
            if i+1 < len(lines): 
                results["HV@NBP (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HIG@Tnbp" in l or "HIG@TNBP(J" in l:
            if i+1 < len(lines): 
                results["HIG@Tnbp (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HIG@Tmax" in l or "HIG@TMAX(J" in l:
            if i+1 < len(lines): 
                results["HIG@Tmax (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1        
        elif "HVAP@NBP" in l or "HVAP@NBP(J" in l:
            if i+1 < len(lines): 
                results["HVAP@NBP (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "PTP@TTP" in l or "PTP@TTP(Kpa)" in l:
            if i+1 < len(lines): 
                results["PTP@TTP (Kpa)"] = extract_last_float(lines[i+1]); i += 1
        elif "Pc_Calc" in l or "Pc_Calc(Kpa)" in l:
            if i+1 < len(lines): 
                results["Pc_Calc (Kpa)"] = extract_last_float(lines[i+1]); i += 1
         #Liquid       
        elif "HL@0C" in l or "HL@0C(J" in l:
            if i+1 < len(lines): 
                results["HL@0C (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@Tmin" in l:
            if i+1 < len(lines): 
                results["HL@Tmin (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@Tmax" in l:
            if i+1 < len(lines): 
                results["HL@Tmax (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@Tnbp" in l:
            if i+1 < len(lines): 
                results["HL@Tnbp (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@Tnmp" in l:
            if i+1 < len(lines): 
                results["HL@Tnmp (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@NBP" in l or "HL@NBP(J" in l:
            if i+1 < len(lines): 
                results["HL@NBP (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HL@NMP" in l or "HL@NMP(J" in l:
            if i+1 < len(lines): 
                results["HL@NMP (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "dHL/dT@Tmin" in l:
            if i+1 < len(lines): 
                results["dHL/dT@Tmin"] = extract_last_float(lines[i+1]); i += 1
        elif "dHL/dT@Tmax" in l:
            if i+1 < len(lines): 
                results["dHL/dT@Tmax"] = extract_last_float(lines[i+1]); i += 1
        elif "dHL/dT@TNBP" in l:
            if i+1 < len(lines): 
                results["dHL/dT@TNBP"] = extract_last_float(lines[i+1]); i += 1
        elif "HVAP@Tnbp" in l:
            if i+1 < len(lines): 
                results["HVAP@Tnbp (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "HVAP@Mid" in l:
            if i+1 < len(lines): 
                results["HVAP@Mid (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1
        elif "LDEN@60F" in l:
            if i+1 < len(lines): 
                results["LDEN@60F (kg/m3)"] = extract_last_float(lines[i+1]); i += 1
        elif "LDEN@25C" in l:
            if i+1 < len(lines): 
                results["LDEN@25C (kg/m3)"] = extract_last_float(lines[i+1]); i += 1
        #Solid
        elif "HS@Tnmp" in l or "HSolid@Tnmp" in l:
            if i+1 < len(lines): 
                results["HS@Tnmp (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1     
        elif "HS@Tmax" in l or "HSolid@Tmax" in l:
            if i+1 < len(lines): 
                results["HS@Tmax (J/kg-mole)"] = extract_last_float(lines[i+1]); i += 1          
        i += 1
    return results

# ---------------------------------------------------
# MAIN WORKFLOW
# ---------------------------------------------------
print(" Loading core properties...")
core_df = pd.read_excel(core_props_excel)
print(f"   Rows: {len(core_df)}, Columns: {len(core_df.columns)}")

# Copy ALL columns (tmin/tmax/eq included)
base_df = core_df.copy()
base_df["TRCID"] = pd.to_numeric(base_df["TRCID"], errors="raise")
base_df.sort_values("TRCID", inplace=True)
# base_df["CASNO"] = base_df["CASNO"].apply(normalize_cas)
base_df["CASNO"] = base_df["CASNO"].apply(lambda x: str(normalize_cas(x)))
print(base_df.loc[base_df["CASNO"].astype(str).str.contains("765"), "CASNO"].head())


# Initialize missing columns
for col, default in [
    ("IC_Temporaryresults", -9999),
    ("Available_in_SIMSCI", "Unknown"),
    ("LCP_Exists", "Unknown"),
    ("ICP_Exists", "Unknown"), 
    ("SCP_Exists", "Unknown"),
    ("HVAP_Exists", "Unknown")
]:
    if col not in base_df.columns:
        base_df[col] = default

SIMSCI_SHEETS = [
    "BioLib",
    "Edlib_SIMSCI",
    "Edlib_PROCESS",
    "Dippr",
    "Dippr_Sponsor"
]

print("SIMSCI lookup (restricted sheets)...")

simsci_set = build_simsci_cas_set(simsci_excel, SIMSCI_SHEETS)

base_df["Available_in_SIMSCI"] = base_df["CASNO"].apply(
    lambda x: "Yes" if x in simsci_set else "No"
)

# JSON flags
print(" JSON flags...")
enthalpy_map = load_enthalpy_flags_from_json(json_folders)

print("Total JSON CAS loaded:", len(enthalpy_map))
print("Example CAS keys:", list(enthalpy_map.keys())[:5])

print("Flags for 765435:", enthalpy_map.get("765435"))
def apply_flags(row):
    cas = row["CASNO"]
    if cas == "765435":
        print("Row CAS 765435, flags:", enthalpy_map.get(cas))
    # flags = enthalpy_map.get(cas, {})
    flags = enthalpy_map.get(str(cas), {})
    row["LCP_Exists"] = flags.get("LCP_Exists", "No")
    row["ICP_Exists"] = flags.get("ICP_Exists", "No")
    row["SCP_Exists"] = flags.get("SCP_Exists", "No")
    row["HVAP_Exists"] = flags.get("HVAP_Exists", "No")
    return row
base_df = base_df.apply(apply_flags, axis=1)

# ---------------------------------------------------
# Initialize sheets - SOURCE DATA FIRST
# ---------------------------------------------------
print("Creating LCP/ICP/SCP/HVAP sheets...")
sheets = {
    "LCP": base_df.copy(),  # LCPtmin/tmax/eq + core props
    "ICP": base_df.copy(),  # ICPtmin/tmax/eq + core props
    "SCP": base_df.copy()   # SCPtmin/tmax/eq + core props
}


# ---------------------------------------------------
# ROUTE PHASE-SPECIFIC COLUMNS TO CORRECT SHEETS ONLY
# ---------------------------------------------------
lcp_cols = ["LCPtmin", "LCPtmax", "LCPEqn"]
scp_cols = ["SCPtmin", "SCPtmax", "SCPEqn"]
icp_cols = ["ICPtmin", "ICPtmax", "ICPEqn"]

# LCP sheet: DROP SCP and ICP columns
for col in scp_cols + icp_cols:
    if col in sheets["LCP"].columns:
        sheets["LCP"].drop(columns=[col], inplace=True)

# SCP sheet: DROP LCP and ICP columns
for col in lcp_cols + icp_cols:
    if col in sheets["SCP"].columns:
        sheets["SCP"].drop(columns=[col], inplace=True)

# ICP sheet: DROP LCP and SCP columns
for col in lcp_cols + scp_cols:
    if col in sheets["ICP"].columns:
        sheets["ICP"].drop(columns=[col], inplace=True)

print("   Phase-specific columns routed:")
print(f"      LCP keeps: {[c for c in lcp_cols if c in sheets['LCP'].columns]}")
print(f"      SCP keeps: {[c for c in scp_cols if c in sheets['SCP'].columns]}")
print(f"      ICP keeps: {[c for c in icp_cols if c in sheets['ICP'].columns]}")

# ---------------------------------------------------
# **CRITICAL: PE1 Processing using ALIAS MATCHING**
# ---------------------------------------------------
for col in ["H_NIST_Trecon", "H_SIMSCI_Trecon"]:
    if col not in sheets["LCP"].columns:
        sheets["LCP"][col] = pd.NA

print("\n PROCESSING PE1 FILES BY ALIAS...")
print("\n PROCESSING PE1 FILES (LEGACY + TRECON)...")
pe1_count = 0
for _, row in base_df.iterrows():

    trcid = int(row["TRCID"])
    trecon = row.get("Trecon (K)")
    comp_name = row["ComponentName"]
    safe_comp = make_safe_name(comp_name)

    H_nist_trecon = pd.NA
    H_sim_trecon  = pd.NA

    # ===============================
    # (C) LEGACY PE1 – ALIAS BASED
    # ===============================
    alias = row["Aliases"]
    if pd.isna(alias):
        continue

    safe_alias = make_safe_name(alias)
    pe1_path_legacy = f"{pe1_dir}\\{safe_alias}_{trcid}_NIST.pe1"
    
#   # Always use TRCID-based alias
#     safe_alias = f"NST{trcid}"

#     pe1_path_legacy = f"{pe1_dir}\\{safe_alias}_{trcid}_NIST.pe1"

#     if not Path(pe1_path_legacy).exists():
#         print(f"[WARNING] PE1 file not found for TRCID {trcid}: {pe1_path_legacy}")
#         continue
    ########

    if Path(pe1_path_legacy).exists():
        parsed = parse_pe1(pe1_path_legacy)

        # for prop, val in parsed.items():

        #         # ---- Special 3 tags go ONLY to SCP ----
        #         if prop in [
        #             "HL@Tnmp (J/kg-mole)",
        #             "dHL/dT@Tmin",
        #             "dHL/dT@Tmax"
        #         ]:
        #             sheets["SCP"].loc[sheets["SCP"]["TRCID"] == trcid, prop] = val

        #         # ---- All other HL go to LCP ----
        #         elif "HL" in prop:
        #             sheets["LCP"].loc[sheets["LCP"]["TRCID"] == trcid, prop] = val

        #         # ---- Vapor phase ----
        #         elif any(x in prop for x in ["HIG", "HVAP", "HV@NBP"]):
        #             sheets["ICP"].loc[sheets["ICP"]["TRCID"] == trcid, prop] = val

        #         # ---- Everything else default to SCP ----
        #         else:
        #             sheets["SCP"].loc[sheets["SCP"]["TRCID"] == trcid, prop] = val

        for prop, val in parsed.items():

            # ---- 1. CONFIG-DRIVEN ROUTING (HIGHEST PRIORITY) ----
            sheet_name = PE1_ROUTING_MAP.get(prop)

            if sheet_name:
                sheets[sheet_name].loc[
                    sheets[sheet_name]["TRCID"] == trcid, prop
                ] = val
                continue

            # ---- 2. SPECIAL SCP CASES ----
            if prop in [
                "HL@Tnmp (J/kg-mole)",
                "dHL/dT@Tmin",
                "dHL/dT@Tmax"
            ]:
                sheets["SCP"].loc[sheets["SCP"]["TRCID"] == trcid, prop] = val

            # ---- 3. HL → LCP ----
            elif "HL" in prop:
                sheets["LCP"].loc[sheets["LCP"]["TRCID"] == trcid, prop] = val

            # ---- 4. VAPOR → ICP ----
            elif any(x in prop for x in ["HIG", "HVAP", "HV@NBP"]):
                sheets["ICP"].loc[sheets["ICP"]["TRCID"] == trcid, prop] = val

            # ---- 5. DEFAULT → SCP (WITH WARNING) ----
            else:
                print(f"[WARNING] Unmapped PE1 property → {prop}")
                sheets["SCP"].loc[
                    sheets["SCP"]["TRCID"] == trcid, prop
                ] = val


        pe1_count += 1

# ---------------------------------------------------
# Calculate HDeparture (ICP only, LAST column)
# ---------------------------------------------------
print(" Calculating HDeparture...")
icp_df = sheets["ICP"]
hv_col = "HV@NBP (J/kg-mole)"
hig_col = "HIG@NBP (J/kg-mole)"

if all(col in icp_df.columns for col in [hv_col, hig_col]):
    icp_df["HDeparture (J/kg-mole)"] = (
        pd.to_numeric(icp_df[hv_col], errors="coerce") - 
        pd.to_numeric(icp_df[hig_col], errors="coerce")
    )
    print(" HDeparture calculated")
else:
    print("HDeparture skipped (missing source columns)")

sheets["ICP"] = icp_df

# ---------------------------------------------------
# Write output - PERFECT column order
# ---------------------------------------------------
print("\n Writing final Excel...")
with pd.ExcelWriter(enthalpy_workflow_excel, engine="openpyxl") as writer:
    for sheet_name, df in sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"   {sheet_name}: {len(df)} rows, {len(df.columns)} columns")

auto_adjust_column_widths(enthalpy_workflow_excel)
print("\n COMPLETE!")
print(f"{enthalpy_workflow_excel}")
print("\nColumn Order:")
print("   LCP:  [Core + LCPtmin/tmax/eq]  [HL@NBP, HL@0C]")
print("   ICP:  [Core + ICPtmin/tmax/eq]  [HIG/HV]  [HDeparture]")
print("   SCP:  [Core + SCPtmin/tmax/eq]  [HS@NMP]")



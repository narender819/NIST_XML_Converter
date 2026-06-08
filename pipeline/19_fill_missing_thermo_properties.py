"""
Script:
    19_fill_missing_thermo_properties.py

Purpose:
    This script fills missing thermodynamic properties in processed
    component JSON files using PE1 workflow results and calculated values.

Functionality:
    - Reads LCP and ICP workflow Excel sheets
    - Processes component JSON files
    - Fills missing properties such as:
        * TC
        * PC
        * ACENTRIC (AF)
        * PTP
        * HVAPNBP
        * SG60F
        * MVOL25C
        * ZNUM
    - Applies validation and correlation range checks
    - Calculates SG60F and MVOL25C values
    - Calculates ZNUM using molecular formula
    - Generates updated JSON files
    - Creates audit and debug reports

Input:
    - LCP/ICP workflow Excel file
    - Processed component JSON files

Output:
    - Updated component JSON files
    - Audit Excel report containing fill status and debug details
"""

import os
import json
import pandas as pd
import re
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

PROCESSED_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ==================================================
# INPUT FILES
# ==================================================
excel_file = (
    PROCESSED_DIR
    / f"14_NIST_Splitsheets_PE1legacy_Hdepart.xlsx"
)

# ==================================================
# JSON INPUT / OUTPUT FOLDERS
# ==================================================
folder_pairs = [

    (
        PROCESSED_DIR
        / "1_components_Inmaster_withsimsciid",

        PROCESSED_DIR
        / "1_components_Inmaster_withsimsciid_fillin"
    ),

    (
        PROCESSED_DIR
        / "3_components_notInmaster_assignedsimsciid",

        PROCESSED_DIR
        / "3_components_notInmaster_assignedsimsciid_fillin"
    )
]

# Create output folders
for _, output_folder in folder_pairs:
    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

# ==================================================
# AUDIT OUTPUT
# ==================================================
audit_output = (
    PROCESSED_DIR
    / "14a_fill_missing_thermo_properties_audit.xlsx"
)

WATER_DENSITY = 997.978
ENABLE_HDN = True

# ==============================
# HELPERS
# ==============================
def normalize_cas(cas):
    if pd.isna(cas):
        return None
    s = str(cas).strip()
    try:
        s = str(int(float(s)))
    except:
        pass
    s = s.replace("-", "")
    s = "".join(c for c in s if c.isdigit())
    return s.lstrip("0")

# ---------- HDN (STRICT ONLY) ----------
def clean_formula(formula):
    return re.sub(r'\[.*?\]', '', str(formula)).strip() if formula else None

def parse_formula(formula):
    elements = re.findall(r'([A-Z][a-z]?)(\d*)', formula)
    counts = {}
    for e, n in elements:
        counts[e] = counts.get(e, 0) + (int(n) if n else 1)
    return counts

def calculate_hdn(formula, debug_msgs):
    if not formula or pd.isna(formula):
        debug_msgs.append("Missing MOLFORMULA")
        return None

    formula = clean_formula(formula)
    counts = parse_formula(formula)

    if not counts:
        debug_msgs.append("Invalid formula parse")
        return None

    C = counts.get("C", 0)
    H = counts.get("H", 0)
    N = counts.get("N", 0)
    X = sum(counts.get(x, 0) for x in ["F", "Cl", "Br", "I"])

    hdn = (2*C + 2 + N - X - H) / 2

    if hdn < 0:
        debug_msgs.append(f"Negative HDN: {hdn}")
        return None

    return int(hdn) if hdn.is_integer() else round(hdn, 3)

# ==============================
# LOAD EXCEL
# ==============================
lcp_df = pd.read_excel(excel_file, sheet_name="LCP")
icp_df = pd.read_excel(excel_file, sheet_name="ICP")

for d in [lcp_df, icp_df]:
    d["TRCID"] = d["TRCID"].astype(str).str.split(".").str[0].str.strip()
    d["CAS_norm"] = d["CASNO"].apply(normalize_cas)
    d.replace(-9999, pd.NA, inplace=True)

lcp_map = {f"{r['TRCID']}_{r['CAS_norm']}": r for _, r in lcp_df.iterrows()}
icp_map = {f"{r['TRCID']}_{r['CAS_norm']}": r for _, r in icp_df.iterrows()}

# ==============================
# PROCESS
# ==============================
audit_rows = []
total = updated = skipped_no_match = errors = 0

for input_folder, output_folder in folder_pairs:
    

    for fname in os.listdir(input_folder):

        if not fname.endswith(".json"):
            continue

        total += 1
        debug_msgs = []
        changed = False

        pc_status = ptp_status = hvap_status = "NOT_FILLED"
        sg60_status = mvol_status = znum_status = "NOT_FILLED"

        try:
            if "_CASNO_" not in fname:
                skipped_no_match += 1
                continue

            trcid = fname.split("_")[0]
            cas = normalize_cas(fname.split("_CASNO_")[1].replace(".json", ""))
            key = f"{trcid}_{cas}"

            lcp_row = lcp_map.get(key)
            icp_row = icp_map.get(key)

            if lcp_row is None and icp_row is None:
                skipped_no_match += 1
                continue

            with open(os.path.join(input_folder, fname)) as f:
                data = json.load(f)

            props = data.get("properties", {})

            def fill(tag, val):
                if (tag not in props or props[tag] is None) and pd.notna(val):
                    props[tag] = float(val)
                    return True
                return False
            # ==========================
            # LCP BASIC PROPERTIES (MISSING FIX)
            # ==========================
            if lcp_row is not None:

                tc_val = lcp_row.get("TC (K)")
                if fill("TC", tc_val):
                    changed = True

                nmp_val = lcp_row.get("NMP (K)")
                if fill("NMP", nmp_val):
                    changed = True

                acentric_val = lcp_row.get("ACENTRIC")
                if fill("ACENTRIC", acentric_val):
                    changed = True

            # ---------- PC ----------
            # pc_val = lcp_row.get("PC (kPa)")
            # if pd.isna(pc_val) and icp_row is not None:
            #     pc_val = icp_row.get("Pc_Calc (Kpa)")

            # if fill("PC", pc_val):
            #     pc_status = "FILLED"; changed = True
            # elif pc_val is None:
            #     pc_status = "SKIPPED"
            
            # ---------- PC (LIGHT VALIDATION) ----------
            # ==========================
            # PC logic (LCP → ICP fallback + garbage filter)
            # ==========================
            pc_val = None

            # ==========================
            # 1. Use LCP directly (trusted)
            # ==========================
            if lcp_row is not None:
                raw_pc = lcp_row.get("PC (kPa)")

                if pd.notna(raw_pc) and raw_pc > 0:
                    pc_val = raw_pc
                else:
                    debug_msgs.append("LCP PC invalid")

            # ==========================
            # 2. Use ICP ONLY if correlation is valid
            # ==========================
            if pc_val is None and icp_row is not None:

                calc_pc = icp_row.get("Pc_Calc (Kpa)")
                tmin = icp_row.get("VPtmin (K)")
                tmax = icp_row.get("VPtmax (K)")
                tc   = lcp_row.get("TC (K)") if lcp_row is not None else None

                #  CRITICAL FIX: check correlation range
                if (
                    pd.notna(calc_pc) and calc_pc > 0 and
                    pd.notna(tc) and pd.notna(tmin) and pd.notna(tmax) and
                    tmin <= tc <= tmax
                ):
                    pc_val = calc_pc
                else:
                    debug_msgs.append("Pc_Calc rejected (out of correlation range)")

            # ==========================
            # 3. Fill
            # ==========================
            if pc_val is not None:
                if fill("PC", pc_val):
                    pc_status = "FILLED"; changed = True
            else:
                pc_status = "SKIPPED"

            # ---------- PTP (LOOSE) ----------
            # ptp_val = icp_row.get("PTP@TTP (Kpa)") if icp_row is not None else None
            # if fill("PTP", ptp_val):
            #     ptp_status = "FILLED"; changed = True
            # elif ptp_val is None:
            #     ptp_status = "SKIPPED"
            
            # ---------- PTP (SAFE LOOSE) ----------
            ptp_val = None

            if icp_row is not None:
                raw_ptp = icp_row.get("PTP@TTP (Kpa)")

                if pd.notna(raw_ptp) and raw_ptp > 1e-6:   # FIX HERE
                    ptp_val = raw_ptp
                else:
                    debug_msgs.append("PTP invalid or too small")

            if fill("PTP", ptp_val):
                ptp_status = "FILLED"; changed = True
            elif ptp_val is None:
                ptp_status = "SKIPPED"

            # # ---------- HVAP (LOOSE) ----------
            # hvap_val = None
            # if icp_row is not None:
            #     raw_hvap = icp_row.get("HVAP@NBP (J/kg-mole)")
            #     if pd.notna(raw_hvap) and raw_hvap != 0:
            #         hvap_val = raw_hvap

            # if fill("HVAPNBP", hvap_val):
            #     hvap_status = "FILLED"; changed = True
            # elif hvap_val is None:
            #     hvap_status = "SKIPPED"


            # ---------- HVAPNBP (SAFE LOOSE + GARBAGE FILTER) ----------
            hvap_val = None

            if icp_row is not None:
                raw_hvap = icp_row.get("HVAP@NBP (J/kg-mole)")

                if pd.notna(raw_hvap) and raw_hvap > 1:   # 🔥 reject tiny garbage
                    hvap_val = raw_hvap
                else:
                    debug_msgs.append("HVAP invalid or too small")

            if fill("HVAPNBP", hvap_val):
                hvap_status = "FILLED"; changed = True
            elif hvap_val is None:
                hvap_status = "SKIPPED"


            # ---------- SG60F (VALIDATED) ----------
            # ---------- SG60F ----------
            sg60 = None

            if lcp_row is not None:

                lden60 = lcp_row.get("LDEN@60F (kg/m3)")
                mw = lcp_row.get("MW")

                temp_60F = 288.76

                ldeqn = lcp_row.get("LDEqn")
                tmin = lcp_row.get("LDtmin (K)")
                tmax = lcp_row.get("LDtmax (K)")

                # =====================================
                # Scenario 1: Correlation missing
                # =====================================
                if ( pd.isna(ldeqn) or str(ldeqn).strip() in ["", "NA", "-9999", "nan"]):

                    sg60 = 0.98765
                    debug_msgs.append("SG60F filled with default value (no correlation)")

                # =====================================
                # Scenario 2: Temperature within range
                # =====================================
                elif (
                    pd.notna(tmin) and
                    pd.notna(tmax) and
                    tmin <= temp_60F <= tmax
                ):

                    if pd.notna(lden60) and pd.notna(mw):
                        sg60 = (lden60 * mw) / WATER_DENSITY

                # =====================================
                # Scenario 3: Extrapolation case
                # =====================================
                else:

                    debug_msgs.append(
                        f"SG60F extrapolation case ({temp_60F} outside {tmin}-{tmax})"
                    )

                    # pass
                    sg60 = 0.98765
                    # Later:
                    # sg60 = 1
            if pd.notna(sg60):
                props["SG60F"] = float(sg60)
                sg60_status = "FILLED"
                changed = True
            elif sg60 is None:
                sg60_status = "SKIPPED"

            # ---------- MVOL25C ----------
            mvol = None

            if lcp_row is not None:

                lden25 = lcp_row.get("LDEN@25C (kg/m3)")
                temp_25C = 298.15

                ldeqn = lcp_row.get("LDEqn")
                tmin = lcp_row.get("LDtmin (K)")
                tmax = lcp_row.get("LDtmax (K)")

                # =====================================
                # Scenario 1: Correlation missing
                # =====================================
                if ( pd.isna(ldeqn) or str(ldeqn).strip() in ["", "NA", "-9999", "nan"]):

                    mvol = 1/0.98765
                    debug_msgs.append("MVOL filled with default value (no correlation)")

                # =====================================
                # Scenario 2: Temperature within range
                # =====================================
                elif (
                    pd.notna(tmin) and
                    pd.notna(tmax) and
                    tmin <= temp_25C <= tmax
                ):

                    if pd.notna(lden25) and lden25 > 0:
                        mvol = 1 / lden25

                # =====================================
                # Scenario 3: Extrapolation case
                # =====================================
                else:

                    debug_msgs.append(
                        f"MVOL extrapolation case ({temp_25C} outside {tmin}-{tmax})"
                    )

                    # pass
                    mvol = 1/0.98765

                    # Later:
                    # mvol = 1

            if pd.notna(mvol):
                props["MVOL25C"] = float(mvol)
                mvol_status = "FILLED"
                changed = True
            elif mvol is None:
                mvol_status = "SKIPPED"

            # ---------- HDN (STRICT) ----------
            if ENABLE_HDN:
                hdn = calculate_hdn(props.get("MOLFORMULA"), debug_msgs)

                if hdn is None:
                    znum_status = "SKIPPED"
                elif fill("ZNUM", hdn):
                    znum_status = "FILLED"; changed = True
                else:
                    znum_status = "ALREADY_PRESENT"
            else:
                znum_status = "IGNORED"

            # ---------- DEBUG ----------
            data["debug_log"] = debug_msgs

            # ---------- ALWAYS SAVE ----------
            with open(os.path.join(output_folder, fname), "w") as f:
                json.dump(data, f, indent=4)

            if changed:
                updated += 1

        except Exception as e:
            errors += 1
            debug_msgs.append(str(e))

        audit_rows.append({
            "TRCID": trcid,
            "CAS": cas,
            "FILE": fname,
            "PC": pc_status,
            "PTP": ptp_status,
            "HVAP": hvap_status,
            "SG60F": sg60_status,
            "MVOL25C": mvol_status,
            "ZNUM": znum_status,
            "DEBUG": " | ".join(debug_msgs)
        })

pd.DataFrame(audit_rows).to_excel(audit_output, index=False)

print(f"Total={total}, Updated={updated}, Errors={errors}")
"""
Script:
    17_run_propeval_for_nist_components.py

Purpose:
    This script generates PropEval input files and executes
    PropEval runs for NIST component thermodynamic data.

Functionality:
    - Reads NIST core thermodynamic property data
    - Generates SLB and TXT input files from templates
    - Replaces template placeholders with component properties
    - Executes PropEval runs for each component
    - Supports parallel processing for faster execution
    - Skips already processed components
    - Tracks execution status and processing summary

Input:
    - NIST thermodynamic property Excel file
    - PropEval TXT and SLB template files

Output:
    - Generated SLB/TXT input files
    - PropEval PE1 output files
    - Processing execution summary
"""

import pandas as pd
import subprocess
import re
from pathlib import Path
from multiprocessing import Pool, cpu_count
from pathlib import Path

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    OUTPUT_DIR,
    NIST_TXT_TEMPLATE,
    NIST_SLB_TEMPLATE,
    PROPEVAL_EXE,
    ensure_directories
)

ensure_directories()

# ==================================================
# OUTPUT DIRECTORIES (runtime)
# ==================================================

PROPEVAL_RUNS_DIR = OUTPUT_DIR / "propeval_runs"
WORK_DIR = PROPEVAL_RUNS_DIR / "NIST"

WORK_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# INPUT FILE
# ==================================================

CORE_EXCEL = PROCESSED_DIR / f"12_NIST_Component_KeyThermoProperties_{RUN_YEAR}_TCPCAF_UPDATED.xlsx"

# ==================================================
# VALIDATION 
# ==================================================

if not CORE_EXCEL.exists():
    raise FileNotFoundError(f"Missing input: {CORE_EXCEL}")

if not PROPEVAL_EXE.exists():
    raise FileNotFoundError(f"Missing executable: {PROPEVAL_EXE}")

if not NIST_TXT_TEMPLATE.exists():
    raise FileNotFoundError(f"Missing template: {NIST_TXT_TEMPLATE}")

if not NIST_SLB_TEMPLATE.exists():
    raise FileNotFoundError(f"Missing template: {NIST_SLB_TEMPLATE}")

# ==================================================
# TEMPLATE ASSIGNMENT
# ==================================================

TXT_TEMPLATE = NIST_TXT_TEMPLATE
SLB_TEMPLATE = NIST_SLB_TEMPLATE

# ==================================================
# DEBUG (optional)
# ==================================================

print("Working directory:", WORK_DIR)
print("Core Excel:", CORE_EXCEL)
print("TXT Template:", TXT_TEMPLATE)
print("SLB Template:", SLB_TEMPLATE)



# ---------------------------------------------------
# UTILITIES
# ---------------------------------------------------
def make_safe_name(name, max_len=30):
    name = name.upper().replace(" ", "_")
    name = re.sub(r"[^A-Z0-9_]", "", name)
    return re.sub(r"_+", "_", name)[:max_len]


def render_template(template, replacements, out_file):
    with open(template, "r", encoding="utf-8") as f:
        content = f.read()
    for k, v in replacements.items():
        content = content.replace(f"{{{k}}}", str(v))
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(content)


def process_component(row):

    try:
        trcid = int(row["TRCID"])
        alias = row["Aliases"]
        comp_library = str(alias).strip()
        safe = make_safe_name(comp_library)

        slb = f"{WORK_DIR}\\{safe}_{trcid}_NIST.slb"
        txt = f"{WORK_DIR}\\{safe}_{trcid}_NIST.txt"
        pe1 = txt.replace(".txt", ".pe1")

        # Skip if already processed
        if Path(pe1).exists():
            return ("SKIPPED", trcid)

        render_template(
            SLB_TEMPLATE,
            {"COMPONENT_NAME": comp_library, "TRCID": trcid},
            slb
        )

        render_template(
            TXT_TEMPLATE,
            {
                "COMPONENT_NAME": comp_library,
                "NBP": row["NBP (K)"],
                "TC": row["TC (K)"],
                "TTP": row["TTP (K)"],
                "NMP": row["NMP (K)"],
                "LCPtmin": row["LCPtmin (K)"],
                "LCPtmax": row["LCPtmax (K)"],
                "ICPtmin": row["ICPtmin (K)"],
                "ICPtmax": row["ICPtmax (K)"],
                "SCPtmin": row["SCPtmin (K)"],
                "SCPtmax": row["SCPtmax (K)"],
                "HFUSIONNMP": row["HFUSIONNMP (kJ/kg-mol)"],
                "ICP_Tnbp": row["ICP_Tnbp (K)"],
                "HVAP_Tmid": row["HVAP_Tmid (K)"],
                "SCP_Tnmp": row["SCP_Tnmp (K)"],
            },
            txt
        )

        subprocess.run(
            [PROPEVAL_EXE, slb, txt],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )

        if not Path(pe1).exists():
            return ("FAILED", trcid)

        return ("DONE", trcid)

    except Exception as e:
        return ("ERROR", str(e))


# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
def main():

    df = pd.read_excel(CORE_EXCEL)
    records = df.to_dict("records")

    # cores = max(cpu_count() - 2, 1)   # leave some CPU free
    cores = min(16, cpu_count()-2)
    print(f"Using {cores} parallel workers")

    with Pool(cores) as pool:
        results = pool.map(process_component, records)

    summary = pd.Series([r[0] for r in results]).value_counts()

    print("\nSummary:")
    print(summary)



if __name__ == "__main__":
    main()
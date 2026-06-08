"""
Script:
    23_generate_phasewise_trecon_propeval_runs.py

Purpose:
    This script generates and executes phasewise Trecon-based
    PropEval runs for both NIST and SIMSCI component data.

Functionality:
    - Reads cleaned phasewise Trecon workflow sheets
    - Merges LCP, ICP, and SCP workflow data
    - Generates PropEval SLB and TXT input files
    - Replaces template placeholders with Trecon temperatures
    - Executes PropEval runs for:
        * NIST component data
        * SIMSCI component data
    - Handles phasewise temperature ranges for:
        * LCP
        * ICP
        * SCP
    - Tracks PropEval execution status

Input:
    - Cleaned phasewise Trecon workflow Excel file
    - NIST and SIMSCI PropEval templates

Output:
    - Generated PropEval SLB/TXT files
    - NIST and SIMSCI PropEval output files
"""

import pandas as pd
import subprocess
import re
from pathlib import Path
from multiprocessing import Pool, cpu_count
import time

start = time.time()



# ==================================================
# CONFIGURATION
# ==================================================
RUN_YEAR = "2025"

BASE_DIR = Path(r"D:\NIST_XML_Converter")

# ==================================================
# PREREQUISITES
# ==================================================
PREREQ_DIR = (
    BASE_DIR
    / "prerequisites"
)

TEMPLATES_DIR = (
    PREREQ_DIR
    / "templates"
)

EXECUTABLES_DIR = (
    PREREQ_DIR
    / "executables"
    / "FlashDriver"
    / "S4MThermo"
    / "Source"
    / "vc17"
    / "x64"
    / "Release"
)

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

PROPEVAL_RUNS_DIR = (
    OUTPUT_DIR
    / "propeval_runs"
)

NIST_WORK_DIR = (
    PROPEVAL_RUNS_DIR
    / "NIST_Trecon"
)

SIMSCI_WORK_DIR = (
    PROPEVAL_RUNS_DIR
    / "SIMSCI_Trecon"
)

NIST_WORK_DIR.mkdir(
    parents=True,
    exist_ok=True
)

SIMSCI_WORK_DIR.mkdir(
    parents=True,
    exist_ok=True
)

# ==================================================
# INPUT FILES
# ==================================================
CORE_EXCEL = (
    PROCESSED_DIR
    / f"17_NIST_With_SIMSCI_Trecon_Clean.xlsx"
)

# ==================================================
# EXECUTABLES
# ==================================================
PROPEVAL_EXE = (
    EXECUTABLES_DIR
    / "PropEval.exe"
)

# ==================================================
# TEMPLATES
# ==================================================
NIST_TEMPLATE_DIR = (
    TEMPLATES_DIR
    / "LibraryCorr_NIST_Trecon"
)

SIMSCI_TEMPLATE_DIR = (
    TEMPLATES_DIR
    / "LibraryCorr_SIMSCI_Trecon"
)

NIST_SLB_TEMPLATE = (
    NIST_TEMPLATE_DIR
    / "{COMPONENT_NAME}_{TRCID}_NIST.slb"
)

NIST_TXT_TEMPLATE = (
    NIST_TEMPLATE_DIR
    / "{COMPONENT_NAME}_{TRCID}_NIST.txt"
)

SIMSCI_SLB_TEMPLATE = (
    SIMSCI_TEMPLATE_DIR
    / "{COMPONENT_NAME}_{LIBID}_SIMSCI.slb"
)

SIMSCI_TXT_TEMPLATE = (
    SIMSCI_TEMPLATE_DIR
    / "{COMPONENT_NAME}_{LIBID}_SIMSCI.txt"
)

# ---------------------------------------------------
# UTILITIES
# ---------------------------------------------------
def make_safe_name(name):
    return re.sub(r"[^A-Z0-9_]", "", str(name).strip().upper().replace(" ", "_"))

def render_template(template_path, out_path, replacements):
    content = Path(template_path).read_text(encoding="utf-8")
    for k, v in replacements.items():
        content = content.replace(f"{{{k}}}", "" if pd.isna(v) else str(v))
    Path(out_path).write_text(content, encoding="utf-8")

# def run_propeval(slb, txt):
#     subprocess.run([PROPEVAL_EXE, slb, txt], check=True)

# def run_propeval(slb, txt):
#     try:
#         subprocess.run([PROPEVAL_EXE, slb, txt], check=True)
#     except subprocess.CalledProcessError as e:
#         print(f"PropEval crashed for: {slb}")
#         print(f"Exit code: {e.returncode}")
#         return False   # indicate failure
    
#     return True        # success

def run_propeval(slb, txt):

    try:

        subprocess.run(

            [
                str(PROPEVAL_EXE),
                str(slb),
                str(txt)
            ],

            cwd=slb.parent,

            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,

            check=True
        )

        return True

    except:

        return False


# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
def main():

    xl = pd.ExcelFile(CORE_EXCEL)

    # ---- Load sheets
    lcp = pd.read_excel(xl, "LCP")
    icp = pd.read_excel(xl, "ICP")
    scp = pd.read_excel(xl, "SCP")

    # ---- Merge into single component table
    df = (
        lcp
        .merge(
            icp,
            on=["TRCID", "CASNO", "Aliases", "LibraryID", "Available_in_SIMSCI"],
            suffixes=("", "_ICP")
        )
        .merge(
            scp,
            on=["TRCID", "CASNO", "Aliases", "LibraryID", "Available_in_SIMSCI"],
            suffixes=("", "_SCP")
        )
    )

    print(f"Total components: {len(df)}")

    for _, row in df.iterrows():

        if not any([
            row.get("LCP_Exists") == "Yes",
            row.get("ICP_Exists") == "Yes",
            row.get("SCP_Exists") == "Yes"
        ]):
            continue

        trcid = int(row["TRCID"])
        alias = make_safe_name(row["Aliases"])
        libid = make_safe_name(row["LibraryID"]) if pd.notna(row["LibraryID"]) else None

        # ---------------------------------------------------
        # TEMPERATURE CONTEXT (single run, multi-phase)
        # ---------------------------------------------------
        repl_common = {
            "COMPONENT_NAME": alias,

            # ---- VAPOR (ICP)
            "ICP_Trecon_SIMSCI": row.get("ICP_Trecon_SIMSCI (K)"),
            "ICP_Trecon_NIST": row.get("ICP_Trecon_NIST (K)"),
            "ICP_Trecon_Remark": row.get("ICP_Trecon_Remark"),
            "IEtmin": row.get("ICPtmin (K)"),
            "IEtmax": row.get("ICPtmax (K)"),

            # ---- LIQUID (LCP)
            "LCP_Trecon_SIMSCI": row.get("LCP_Trecon_SIMSCI (K)"),
            "LCP_Trecon_NIST": row.get("LCP_Trecon_NIST (K)"),
            "LCP_Trecon_Remark": row.get("LCP_Trecon_Remark"),
            "LEtmin": row.get("LCPtmin (K)"),
            "LEtmax": row.get("LCPtmax (K)"),

            # ---- SOLID (SCP)
            "SCP_Trecon_SIMSCI": row.get("SCP_Trecon_SIMSCI (K)"),
            "SCP_Trecon_NIST": row.get("SCP_Trecon_NIST (K)"),
            "SCP_Trecon_Remark": row.get("SCP_Trecon_Remark"),
            "SEtmin": row.get("SCPtmin (K)"),
            "SEtmax": row.get("SCPtmax (K)"),
        }

        # ===================================================
        # NIST PROPEVAL
        # ===================================================
        slb_nist = Path(NIST_WORK_DIR) / f"{alias}_{trcid}_NIST.slb"
        txt_nist = Path(NIST_WORK_DIR) / f"{alias}_{trcid}_NIST.txt"

        # render_template(
        #     NIST_SLB_TEMPLATE,
        #     slb_nist,
        #     {"COMPONENT_NAME": alias, "TRCID": trcid}
        # )

        # render_template(
        #     NIST_TXT_TEMPLATE,
        #     txt_nist,
        #     repl_common
        # )

        # run_propeval(str(slb_nist), str(txt_nist))

        pe1_nist = txt_nist.with_suffix(".pe1")

        if not pe1_nist.exists():

            render_template(
                NIST_SLB_TEMPLATE,
                slb_nist,
                {"COMPONENT_NAME": alias, "TRCID": trcid}
            )

            render_template(
                NIST_TXT_TEMPLATE,
                txt_nist,
                repl_common
            )

            run_propeval(
                slb_nist,
                txt_nist
            )

        # ===================================================
        # SIMSCI PROPEVAL
        # ===================================================
        if row.get("Available_in_SIMSCI") == "Yes" and libid:

            slb_sim = Path(SIMSCI_WORK_DIR) / f"{libid}_{libid}_SIMSCI.slb"
            txt_sim = Path(SIMSCI_WORK_DIR) / f"{libid}_{libid}_SIMSCI.txt"

            # render_template(
            #     SIMSCI_SLB_TEMPLATE,
            #     slb_sim,
            #     {"COMPONENT_NAME": libid, "LIBID": libid}
            # )

            repl_sim = repl_common.copy()
            repl_sim["COMPONENT_NAME"] = libid

            # overwrite temps with SIMSCI ranges if present
            repl_sim.update({
                "IEtmin": row.get("IEtmin_simsci"),
                "IEtmax": row.get("IEtmax_simsci"),
                "LEtmin": row.get("LEtmin_simsci"),
                "LEtmax": row.get("LEtmax_simsci"),
                "SEtmin": row.get("SEtmin_simsci"),
                "SEtmax": row.get("SEtmax_simsci"),
            })

            # render_template(
            #     SIMSCI_TXT_TEMPLATE,
            #     txt_sim,
            #     repl_sim
            # )

            # run_propeval(str(slb_sim), str(txt_sim))

            pe1_sim = txt_sim.with_suffix(".pe1")

            if not pe1_sim.exists():

                render_template(
                    SIMSCI_SLB_TEMPLATE,
                    slb_sim,
                    {
                        "COMPONENT_NAME": libid,
                        "LIBID": libid
                    }
                )

                render_template(
                    SIMSCI_TXT_TEMPLATE,
                    txt_sim,
                    repl_sim
                )

                run_propeval(
                    slb_sim,
                    txt_sim
                )

        if trcid % 100 == 0:
            print(f"Completed TRCID {trcid}")


# ---------------------------------------------------
if __name__ == "__main__":
    main()

end = time.time()

elapsed = end - start

print(
    f"\nTotal runtime: "
    f"{elapsed:.2f} sec "
    f"({elapsed/60:.2f} min)"
)
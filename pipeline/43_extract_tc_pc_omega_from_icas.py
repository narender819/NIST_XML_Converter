import os
from pathlib import Path
import pandas as pd
from config import (
    TC_PC_SMILES_FILE,
    ICAS_PROCESSED_OUTPUTS_DIR,
    TC_PC_AF_EXTRACTED_FILE
)
# ---------------------------------------------------
# PATHS
# ---------------------------------------------------
EXCEL_PATH = TC_PC_SMILES_FILE

TXT_DIR = ICAS_PROCESSED_OUTPUTS_DIR

OUTPUT_EXCEL = TC_PC_AF_EXTRACTED_FILE

# EXCEL_PATH = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\3_NIST_TC_PC_SMILES_Merged_sep.xlsx")
# TXT_DIR = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\icas_processed_outputs")
# # OUTPUT_EXCEL = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\4_NIST_TC_PC_OMEGA_From_ICAS.xlsx")
# OUTPUT_EXCEL = Path(r"D:\NIST_XML_Converter\prerequisites\excel_inputs\7_TC_PC_AF_extracted.xlsx")

SHEET_NAME = "SMILES_FOUND"
os.makedirs(OUTPUT_EXCEL.parent, exist_ok=True)

# ---------------------------------------------------
# HELPERS
# ---------------------------------------------------
def get_main_fragment(smiles):
    parts = smiles.split(".")
    return max(parts, key=len)

def is_smiles_line(line):
    return (
        line
        and " " not in line
        and "invalid" not in line.lower()
        and not line.startswith((
            "Property", "Primary", "Secondary",
            "WARNING", "Compound", "Temp",
            "-", "_"
        ))
    )

def normalize_smiles(s):
    return str(s).strip().replace(" ", "")

def extract_numeric(parts):
    """
    Extract first valid float from a list of tokens.
    Ignores N/A, ******, etc.
    """
    for p in parts:
        try:
            val = float(p)
            return val
        except:
            continue
    return None

# ---------------------------------------------------
# EXTRACTION (BLOCK-BASED)
# ---------------------------------------------------
def extract_propred_data(txt_dir: Path):
    tc_results = {}
    pc_results = {}
    omega_results = {}
    invalid_results = set()
    processed_smiles = set()

    for txt_file in txt_dir.rglob("*.txt"):
        print(f"Processing: {txt_file.name}")

        try:
            lines = txt_file.read_text(encoding="utf-8", errors="ignore").splitlines()
        except:
            continue

        n = len(lines)
        i = 0

        while i < n:
            line = lines[i].strip()

            # ---------------------------
            # 1. INVALID HANDLING
            # ---------------------------
            if "invalid" in line.lower():
                raw = normalize_smiles(line.replace("Invalid", "").strip())
                smiles = get_main_fragment(raw)

                if smiles:
                    invalid_results.add(smiles)

                i += 1
                continue

            # ---------------------------
            # 2. VALID BLOCK START
            # ---------------------------
            if is_smiles_line(line):

                raw_smiles = normalize_smiles(line)
                compound_smiles = None

                tc_val = None
                pc_val = None
                omega_val = None   #  FIXED

                in_primary = False

                i += 1

                while i < n:
                    l = lines[i].strip()

                    # ---------------------------
                    # BLOCK END
                    # ---------------------------
                    if "Secondary Properties" in l:
                        break

                    # ---------------------------
                    # COMPOUND SMILES
                    # ---------------------------
                    if "Compound Smiles" in l:
                        try:
                            compound_smiles = normalize_smiles(l.split(":")[1])
                        except:
                            pass

                    # ---------------------------
                    # PRIMARY SECTION START
                    # ---------------------------
                    if "Primary Properties" in l:
                        in_primary = True
                        i += 1
                        continue

                    # ---------------------------
                    # EXTRACT PROPERTIES
                    # ---------------------------
                    if in_primary:
                        l2 = l.strip()
                        parts = l2.split()

                        if l2.startswith("Tc"):
                            tc_val = extract_numeric(parts)

                        elif l2.startswith("Pc"):
                            pc_val = extract_numeric(parts)

                        elif l2.startswith("omega"):
                            omega_val = extract_numeric(parts)

                    i += 1

                # ---------------------------
                # FINAL SMILES SET
                # ---------------------------
                smiles_set = {raw_smiles}

                if compound_smiles:
                    smiles_set.add(compound_smiles)

                if raw_smiles not in invalid_results:
                    processed_smiles.update(smiles_set)

                # ---------------------------
                # STORE RESULTS
                # ---------------------------
                for sm in smiles_set:

                    if tc_val is not None:
                        tc_results[sm] = tc_val

                    if pc_val is not None:
                        pc_results[sm] = pc_val

                    if omega_val is not None:
                        omega_results[sm] = omega_val

                continue

            i += 1

    return tc_results, pc_results, invalid_results, omega_results, processed_smiles

# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
def main():
    print("FINAL BLOCK-BASED EXTRACTOR (TC / PC / OMEGA)")

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    df["SMILES_NORM"] = df["SMILES"].apply(lambda x: get_main_fragment(str(x).strip()))

    tc_results, pc_results, invalid_results, omega_results, processed_smiles = extract_propred_data(TXT_DIR)

    # Initialize columns
    df["TC"] = "NA"
    df["PC"] = "NA"
    df["OMEGA"] = "NA"
    df["STATUS"] = "NOT_FOUND"

    # -------------------------------
    # PRIORITY 1: INVALID
    # -------------------------------
    for smiles in invalid_results:
        mask = df["SMILES_NORM"] == smiles
        df.loc[mask, "STATUS"] = "INVALID_SMILES"
        df.loc[mask, ["TC", "PC", "OMEGA"]] = "NA"

    # -------------------------------
    # PRIORITY 2: FOUND
    # -------------------------------
    for smiles, val in tc_results.items():
        mask = df["SMILES_NORM"] == smiles
        df.loc[mask, "TC"] = val
        df.loc[mask, "STATUS"] = "FOUND"

    for smiles, val in pc_results.items():
        mask = df["SMILES_NORM"] == smiles
        df.loc[mask, "PC"] = val
        df.loc[(mask) & (df["STATUS"] != "INVALID_SMILES"), "STATUS"] = "FOUND"

    for smiles, val in omega_results.items():
        mask = df["SMILES_NORM"] == smiles
        df.loc[mask, "OMEGA"] = val
        df.loc[(mask) & (df["STATUS"] != "INVALID_SMILES"), "STATUS"] = "FOUND"

    # -------------------------------
    # PRIORITY 3: PROPERTY_MISSING
    # -------------------------------
    mask_pm = (
        df["SMILES_NORM"].isin(processed_smiles) &
        (df["STATUS"] == "NOT_FOUND")
    )
    df.loc[mask_pm, "STATUS"] = "PROPERTY_MISSING"

    # -------------------------------
    # SUMMARY
    # -------------------------------
    summary_df = pd.DataFrame({
        "Metric": [
            "Total", "FOUND", "INVALID_SMILES",
            "PROPERTY_MISSING", "NOT_FOUND",
            "TC_Found", "PC_Found", "OMEGA_Found"
        ],
        "Count": [
            len(df),
            (df["STATUS"] == "FOUND").sum(),
            (df["STATUS"] == "INVALID_SMILES").sum(),
            (df["STATUS"] == "PROPERTY_MISSING").sum(),
            (df["STATUS"] == "NOT_FOUND").sum(),
            (df["TC"] != "NA").sum(),
            (df["PC"] != "NA").sum(),
            (df["OMEGA"] != "NA").sum()
        ]
    })

    # -------------------------------
    # SAVE
    # -------------------------------
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DATA")
        summary_df.to_excel(writer, index=False, sheet_name="SUMMARY")

    print("\nDONE")
    print(summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
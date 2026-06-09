"""
Script:
    16_build_nist_simsci_component_availability.py

Purpose:
    This script identifies NIST components already available
    in SIMSCI libraries and generates a consolidated availability report.

Functionality:
    - Reads NIST thermodynamic property data
    - Reads SIMSCI library extraction data
    - Normalizes CAS numbers for matching
    - Builds prioritized SIMSCI component lookup mappings
    - Matches NIST components against SIMSCI libraries
    - Extracts existing SIMSCI IDs and correlation ranges
    - Generates a consolidated NIST-SIMSCI availability report

Input:
    - NIST thermodynamic property Excel file
    - SIMSCI library extraction Excel file

Output:
    Excel report containing NIST components available in SIMSCI libraries
"""

import pandas as pd

from pathlib import Path

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

NIST_EXCEL = PROCESSED_DIR / f"12_NIST_Component_KeyThermoProperties_{RUN_YEAR}_TCPCAF_UPDATED.xlsx"

SIMSCI_EXCEL = PROCESSED_DIR / "2_Libraries_XML_Component_Extract.xlsx"

OUTPUT_EXCEL = PROCESSED_DIR / f"13_NIST_Components_Available_in_SIMSCI_{RUN_YEAR}.xlsx"

SIMSCI_SHEETS_PRIORITY = [
    "Edlib_SIMSCI",
    "Edlib_PROCESS",
    "BioLib",
    "Dippr",
]

# ---------------------------------------------------
# HELPERS
# ---------------------------------------------------
def normalize_cas(cas):
    if pd.isna(cas):
        return None
    return "".join(c for c in str(cas).strip() if c.isdigit())


def find_column(df, keywords):
    for col in df.columns:
        col_upper = str(col).upper()
        for key in keywords:
            if key in col_upper:
                return col
    return None


# ---------------------------------------------------
# MAIN
# ---------------------------------------------------
def build_nist_simsci_availability():

    print("Loading NIST Excel...")
    nist_df = pd.read_excel(NIST_EXCEL)
    print(f"NIST rows: {len(nist_df)}")

    print("Loading SIMSCI sheets...")
    simsci_sheets = pd.read_excel(SIMSCI_EXCEL, sheet_name=None)

    # ---------------------------------------------------
    # BUILD FAST LOOKUP DICTIONARY (PRIORITY BASED)
    # ---------------------------------------------------
    simsci_lookup = {}

    for sheet in SIMSCI_SHEETS_PRIORITY:

        if sheet not in simsci_sheets:
            continue

        sim_df = simsci_sheets[sheet]

        cas_col = find_column(sim_df, ["CAS"])
        letmin_col = find_column(sim_df, ["LETMIN"])
        letmax_col = find_column(sim_df, ["LETMAX"])
        ietmin_col = find_column(sim_df, ["IETMIN"])
        ietmax_col = find_column(sim_df, ["IETMAX"])
        setmin_col = find_column(sim_df, ["SETMIN"])
        setmax_col = find_column(sim_df, ["SETMAX"])
        id_col = find_column(sim_df, ["SIMSCI", "ID"])
        # lib_col = find_column(sim_df, ["LIBRARYID"])
        lib_col = find_column(

    sim_df,

    [

        "LIBRARYID",

        "LIBRARY",

        "LIB_ID",

        "LIB"

    ]

)
        print(sheet,sim_df.columns.tolist())
        print(f"{sheet}:", "Library column =", lib_col)

        

        if cas_col is None:
            continue

        for _, sim_row in sim_df.iterrows():

            sim_cas = normalize_cas(sim_row[cas_col])
            if not sim_cas:
                continue

            # Only store if not already captured by higher priority sheet
            if sim_cas not in simsci_lookup:

                simsci_lookup[sim_cas] = {
                    "SIMSCI ID": sim_row[id_col] if id_col else None,
                    "LibraryID": sim_row[lib_col] if lib_col else None,
                    "Source_Sheet": sheet,
                    "LEtmin_simsci": sim_row.get(letmin_col),
                    "LEtmax_simsci": sim_row.get(letmax_col),
                    "IEtmin_simsci": sim_row.get(ietmin_col),
                    "IEtmax_simsci": sim_row.get(ietmax_col),
                    "SEtmin_simsci": sim_row.get(setmin_col),
                    "SEtmax_simsci": sim_row.get(setmax_col),
                }

    print("SIMSCI lookup built.")
    print(f"Total unique CAS in lookup: {len(simsci_lookup)}")

    # # ---------------------------------------------------
    # # MATCH NIST USING FAST LOOKUP
    # # ---------------------------------------------------
    # rows = []
    # matched_count = 0

    # for _, nist_row in nist_df.iterrows():

    #     nist_cas = normalize_cas(nist_row["CASNO"])
    #     if not nist_cas:
    #         continue

    #     if nist_cas in simsci_lookup:

    #         sim_data = simsci_lookup[nist_cas]

    #         rows.append({
    #             "NIST ID": nist_row["TRCID"],
    #             "CASNO": nist_cas,
    #             "Name": nist_row["ComponentName"],
    #             **sim_data
    #         })

    #         matched_count += 1
    

        # ---------------------------------------------------
    # MATCH NIST USING FAST LOOKUP
    # ---------------------------------------------------
    rows = []

    matched_count = 0


    for _, nist_row in nist_df.iterrows():

        row = (
            nist_row
            .to_dict()
        )


        nist_cas = normalize_cas(

            nist_row[
                "CASNO"
            ]

        )


        if (

            nist_cas

            and

            nist_cas in simsci_lookup

        ):

            sim_data = (

                simsci_lookup[
                    nist_cas
                ]

            )


            row.update(sim_data)       


            row[
                "Available_in_SIMSCI"
            ] = "Yes"


            matched_count += 1


        else:

            row[
                "Available_in_SIMSCI"
            ] = "No"


            row[
                "SIMSCI ID"
            ] = None


            row[
                "LibraryID"
            ] = None


            row[
                "Source_Sheet"
            ] = None



        rows.append(
            row
        )

    # ---------------------------------------------------
    # WRITE OUTPUT
    # ---------------------------------------------------
    out_df = pd.DataFrame(rows)
    out_df.insert(0, "RowID", range(1, len(out_df) + 1))
    out_df.to_excel(OUTPUT_EXCEL, index=False)

    print("\nScript DONE")
    print("Matched NIST components:", matched_count)
    print("Rows written:", len(out_df))
    print("Output:", OUTPUT_EXCEL)


if __name__ == "__main__":
    build_nist_simsci_availability()
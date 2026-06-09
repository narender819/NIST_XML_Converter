import pandas as pd
import numpy as np
from pathlib import Path


from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# INPUT FILES
# ==================================================

NIST_EXCEL = PROCESSED_DIR / "14_NIST_Splitsheets_PE1legacy_Hdepart.xlsx"

SIMSCI_EXCEL = PROCESSED_DIR / f"13_NIST_Components_Available_in_SIMSCI_{RUN_YEAR}.xlsx"

# ==================================================
# OUTPUT FILE
# ==================================================

OUTPUT_EXCEL = PROCESSED_DIR / "15_NIST_Phasewise_With_SIMSCI_Metadata.xlsx"


def normalize_cas(cas):

    if pd.isna(cas):
        return None

    cas = str(cas).strip()

    if cas.endswith(".0"):
        cas = cas[:-2]

    return "".join(
        c
        for c in cas
        if c.isdigit()
    )


def merge_nist_simsci():

    simsci = pd.read_excel(
        SIMSCI_EXCEL
    )

    keep_cols = [

        "CASNO",

        "SIMSCI ID",

        "LibraryID",

        "Source_Sheet",

        "LEtmin_simsci",
        "LEtmax_simsci",

        "IEtmin_simsci",
        "IEtmax_simsci",

        "SEtmin_simsci",
        "SEtmax_simsci"

    ]

    keep_cols = [
        c
        for c in keep_cols
        if c in simsci.columns
    ]

    simsci = simsci[
        keep_cols
    ]

    simsci["CASNO"] = (
        simsci["CASNO"]
        .apply(normalize_cas)
    )

    # remove duplicate CAS
    simsci = (
        simsci
        .drop_duplicates(
            subset=["CASNO"],
            keep="first"
        )
    )

    print(
    "SIMSCI components with valid ID:",
        simsci["SIMSCI ID"]
        .notna()
        .sum()
    )

    print("Total SIMSCI components:", len(simsci))


    nist_sheets = pd.read_excel(
        NIST_EXCEL,
        sheet_name=None
    )

    output = {}


    for sheet, df in nist_sheets.items():

        print(
            "\nProcessing:",
            sheet
        )

        df["CASNO"] = (
            df["CASNO"]
            .apply(
                normalize_cas
            )
        )
        # print("Unique NIST CAS:", df["CASNO"].nunique())

        # print(
        #     "Unique SIMSCI CAS:",
        #     simsci["CASNO"].nunique()
        # )

        # print(
        #     "Intersection:",
        #     len(
        #         set(df["CASNO"])
        #         &
        #         set(simsci["CASNO"])
        #     )
        # )

        merged = df.merge(

            simsci,

            on="CASNO",

            how="left",

            validate="many_to_one"

        )


        merged[
            "Available_in_SIMSCI"
        ] = np.where(

            merged[
                "SIMSCI ID"
            ].notna(),

            "Yes",

            "No"

        )
        print("Total Yes:",(merged["Available_in_SIMSCI"]=="Yes").sum())

        print("Total No:",(merged["Available_in_SIMSCI"]=="No").sum())

        print( "Input rows:", len(df))

        print(

            "Output rows:",
            len(merged)

        )

        print(

            "Matched:",

            merged[
                "Available_in_SIMSCI"
            ].eq(
                "Yes"
            ).sum()

        )


        output[
            sheet
        ] = merged


    with pd.ExcelWriter(
        OUTPUT_EXCEL,
        engine="openpyxl"
    ) as writer:

        for sheet, df in output.items():

            df.to_excel(

                writer,

                sheet_name=sheet,

                index=False

            )


    print(
        "\nCompleted"
    )


if __name__ == "__main__":

    merge_nist_simsci()
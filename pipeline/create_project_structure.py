"""
Script:
create_project_structure.py

Purpose:
Creates the standard folder structure required for
the NIST Library Generation workflow.

Usage:
python pipeline/create_project_structure.py
"""

from pathlib import Path

# --------------------------------------------------
# Project Root
# --------------------------------------------------

PROJECT_ROOT = Path.cwd()

# --------------------------------------------------
# User Input
# --------------------------------------------------

YEAR = input(
    "\nEnter library refresh year "
    "(Example: 2025): "
).strip()

if not YEAR:
    print("\n Invalid year provided.")
    exit()

# --------------------------------------------------
# Folder Structure
# --------------------------------------------------

FOLDERS = [

    # Root Folders
    "batfiles",
    "db",
    "Documents",
    "logs",
    "output",
    "pipeline",
    "prerequisites",

    # Database Structure
    f"db/{YEAR}",

    # Output Structure
    f"output/{YEAR}",
    f"output/{YEAR}/json",
    f"output/{YEAR}/processed",
    f"output/{YEAR}/xml",
    f"output/{YEAR}/libraries",
    f"output/{YEAR}/logs",
    f"output/{YEAR}/propeval_runs",
    f"output/{YEAR}/smiles",
    f"output/{YEAR}/temp",
    f"output/{YEAR}/skipped",

    # Prerequisites Structure
    "prerequisites/excel_inputs",
    "prerequisites/executables",
    "prerequisites/intermediate_library_generation",
    "prerequisites/libraries",
    "prerequisites/templates",

    # ICAS Structure
    "prerequisites/Smiles_Prep_ICAS",
    "prerequisites/Smiles_Prep_ICAS/icas_smiles_batches",
    "prerequisites/Smiles_Prep_ICAS/icas_processed_outputs",
]

# --------------------------------------------------
# Create Folders
# --------------------------------------------------

created = []
existing = []
failed = []

print("\n" + "=" * 70)
print("NIST Project Folder Structure Initialization")
print("=" * 70)

print(f"\nSelected Release Year : {YEAR}")

for folder in FOLDERS:

    folder_path = PROJECT_ROOT / folder

    try:

        if not folder_path.exists():

            folder_path.mkdir(
                parents=True,
                exist_ok=True
            )

            created.append(folder)

            print(f"[CREATED] {folder}")

        else:

            existing.append(folder)

            print(f"[EXISTS ] {folder}")

    except Exception as e:

        failed.append((folder, str(e)))

        print(f"[ERROR  ] {folder}")
        print(f"          {e}")

# --------------------------------------------------
# Summary
# --------------------------------------------------

print("\n" + "=" * 70)
print("Folder Structure Creation Summary")
print("=" * 70)

print(f"\nTotal Configured Folders : {len(FOLDERS)}")
print(f"Created Folders          : {len(created)}")
print(f"Existing Folders         : {len(existing)}")
print(f"Failed Folders           : {len(failed)}")

# --------------------------------------------------
# Failed Folders
# --------------------------------------------------

if failed:

    print("\nFailed Folder Creation Details:")

    for folder, error in failed:

        print(f"  - {folder}")
        print(f"    Error: {error}")

# --------------------------------------------------
# Completion Message
# --------------------------------------------------

print("\n" + "=" * 70)

if len(failed) == 0:

    print(
        "Folder structure initialization "
        "completed successfully."
    )

else:

    print(
        "Folder structure initialization "
        "completed with errors."
    )

print("=" * 70)
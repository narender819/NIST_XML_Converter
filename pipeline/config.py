from pathlib import Path


# ==================================
# RUN SETTINGS
# ==================================

RUN_YEAR = "2025"
TEST_MODE = False
MAX_COMPONENTS = None
N_WORKERS = 8
CHUNKSIZE = 50


# ==================================
# PROJECT ROOT
# ==================================

BASE_DIR = Path(__file__).resolve().parent.parent

PREREQ_DIR = BASE_DIR / "prerequisites"

OUTPUT_DIR = BASE_DIR / "output" / RUN_YEAR


# ==================================
# COMMON OUTPUT PATHS
# ==================================

PROCESSED_DIR = OUTPUT_DIR / "processed" / "full_library"

SMILES_DIR = OUTPUT_DIR / "smiles"

XML_DIR = OUTPUT_DIR / "xml"

LOG_DIR = OUTPUT_DIR / "logs"

TEMP_DIR = OUTPUT_DIR / "temp"

SKIPPED_DIR = OUTPUT_DIR / "skipped"

LB1_DIR = OUTPUT_DIR / "lb1"

LB2_DIR = OUTPUT_DIR / "lb2"

LIBRARY_OUTPUT_DIR = OUTPUT_DIR / "libraries"


# ==================================
# XML WORKFLOW
# ==================================

XML_LIBRARY_DIR = XML_DIR / "Libraryfiles_NIST"


# ==================================
# INTERMEDIATE LIBRARY WORKFLOW
# ==================================

INTERMEDIATE_LIBRARY_DIR = (
    PREREQ_DIR / "intermediate_library_generation"
)

INTERMEDIATE_XML_DIR = (
    INTERMEDIATE_LIBRARY_DIR / "xml"
)

INTERMEDIATE_LIBRARY_OUTPUT_DIR = (
    INTERMEDIATE_LIBRARY_DIR / "library"
)


# ==================================
# TEMPLATE DIRECTORIES
# ==================================

TEMPLATE_NIST_DIR = (
    PREREQ_DIR / "templates" / "LibraryCorr_NIST"
)

TEMPLATE_SIMSCI_TRECON_DIR = (
    PREREQ_DIR / "templates" / "LibraryCorr_SIMSCI_Trecon"
)

TEMPLATE_NIST_TRECON_DIR = (
    PREREQ_DIR / "templates" / "LibraryCorr_NIST_Trecon"
)


# ==================================
# NIST TEMPLATES
# ==================================

NIST_TXT_TEMPLATE = TEMPLATE_NIST_DIR / "__NIST.txt"

NIST_SLB_TEMPLATE = TEMPLATE_NIST_DIR / "__NIST.slb"


# ==================================
# SIMSCI TRECON TEMPLATES
# ==================================

SIMSCI_TRECON_TXT_TEMPLATE = (
    TEMPLATE_SIMSCI_TRECON_DIR / "__SIMSCI.txt"
)

SIMSCI_TRECON_SLB_TEMPLATE = (
    TEMPLATE_SIMSCI_TRECON_DIR / "__SIMSCI.slb"
)


# ==================================
# NIST TRECON TEMPLATES
# ==================================

NIST_TRECON_TXT_TEMPLATE = (
    TEMPLATE_NIST_TRECON_DIR / "__NIST.txt"
)

NIST_TRECON_SLB_TEMPLATE = (
    TEMPLATE_NIST_TRECON_DIR / "__NIST.slb"
)


# ==================================
# PROPEVAL
# ==================================

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

PROPEVAL_EXE = EXECUTABLES_DIR / "PropEval.exe"


# ==================================
# TC / PC RECOVERY WORKFLOW (39-43)
# ==================================

SMILES_PREP_DIR = (
    PREREQ_DIR / "Smiles_Prep_ICAS"
)

SMILES_MASTER_FILE = (
    SMILES_DIR / f"1_compounds_smiles_{RUN_YEAR}.xlsx"
)

INMASTER_COMPONENTS_DIR = (
    PROCESSED_DIR / "1_components_Inmaster_withsimsciid"
)

NOTINMASTER_COMPONENTS_DIR = (
    PROCESSED_DIR / "3_components_notInmaster_assignedsimsciid"
)

COMPONENT_JSON_DIRS = [
    INMASTER_COMPONENTS_DIR,
    NOTINMASTER_COMPONENTS_DIR
]

TC_PC_MISSING_REPORT = (
    SMILES_PREP_DIR / "1_TC_PC_Missing_Report.xlsx"
)

TC_PC_MERGED_FILE = (
    SMILES_PREP_DIR / "2_NIST_TC_PC_Merged.xlsx"
)

TC_PC_SMILES_FILE = (
    SMILES_PREP_DIR / "3_NIST_TC_PC_SMILES_Merged_sep.xlsx"
)

ICAS_SMILES_BATCHES_DIR = (
    SMILES_PREP_DIR / "icas_smiles_batches"
)

ICAS_PROCESSED_OUTPUTS_DIR = (
    SMILES_PREP_DIR / "icas_processed_outputs"
)

TC_PC_OMEGA_FROM_ICAS_FILE = (
    SMILES_PREP_DIR / "4_NIST_TC_PC_OMEGA_From_ICAS.xlsx"
)

TC_PC_AF_EXTRACTED_FILE = (
    PREREQ_DIR
    / "excel_inputs"
    / "7_TC_PC_AF_extracted.xlsx"
)


# ==================================
# HELPERS
# ==================================

def ensure_directories():

    for path in [
        OUTPUT_DIR,
        PROCESSED_DIR,
        XML_DIR,
        XML_LIBRARY_DIR,
        LOG_DIR,
        TEMP_DIR,
        SKIPPED_DIR,
        LIBRARY_OUTPUT_DIR,
        INTERMEDIATE_XML_DIR,
        INTERMEDIATE_LIBRARY_OUTPUT_DIR,
        ICAS_SMILES_BATCHES_DIR,
        ICAS_PROCESSED_OUTPUTS_DIR
    ]:
        path.mkdir(parents=True, exist_ok=True)
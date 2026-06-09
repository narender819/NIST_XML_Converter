# I have an excel sheet D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\NIST_TC_PC_SMILES_Merged_sep.xlsx which comtain sheet name "SMILES_FOUND" and sample data is like TRCID	TC_Status	PC_Status	CASRN	SMILES
#5	PROPERTY_MISSING	PROPERTY_MISSING	2004037	Cc1ncnc2[nH]cnc12 now my task to copy those smiles in to  txt file automatically satrting 1 to 10000 in one txt file & next 1001 to 2000 in next file & so on untli the list completes(Total smiles list: 11712). name od the next file is like 1_Smiles_TC_PC_Missing_NIST_1to1000_raw.txt 1_Smiles_TC_PC_Missing_NIST_1001to2000_raw.txt & so on. path of txt files is D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\raw1

import pandas as pd
from pathlib import Path

from config import TC_PC_SMILES_FILE, ICAS_SMILES_BATCHES_DIR
# ---------------------------------------------------
# PATHS
EXCEL_PATH = TC_PC_SMILES_FILE
OUTPUT_DIR = ICAS_SMILES_BATCHES_DIR


# EXCEL_PATH = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\3_NIST_TC_PC_SMILES_Merged_sep.xlsx")
# OUTPUT_DIR = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\icas_smiles_batches") 




# ---------------------------------------------------
# ---------------------------------------------------
# LOAD EXCEL
df = pd.read_excel(EXCEL_PATH, sheet_name="SMILES_FOUND")
# ---------------------------------------------------
# EXTRACT SMILES
smiles_list = df["SMILES"].tolist()
# ---------------------------------------------------
# FUNCTION TO WRITE SMILES TO TXT FILES
def write_smiles_to_txt(smiles, start_idx, end_idx):
    file_name = f"{start_idx}_Smiles_TC_PC_Missing_NIST_{start_idx}to{end_idx}_raw.txt"
    file_path = OUTPUT_DIR / file_name
    with open(file_path, "w") as f:
        for smile in smiles:
            f.write(smile + "\n")

# ---------------------------------------------------
# SPLIT SMILES INTO CHUNKS AND WRITE TO FILES
chunk_size = 1000
for i in range(0, len(smiles_list), chunk_size):
    start_idx = i + 1
    end_idx = min(i + chunk_size, len(smiles_list))
    write_smiles_to_txt(smiles_list[i:end_idx], start_idx, end_idx)
print("SMILES have been successfully written to text files.")


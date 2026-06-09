from pathlib import Path
import json
import pandas as pd

from config import COMPONENT_JSON_DIRS, TC_PC_MISSING_REPORT

INPUT_FOLDERS = COMPONENT_JSON_DIRS

OUTPUT_FILE = TC_PC_MISSING_REPORT

# INPUT_FOLDERS = [
#     Path(r"D:\NIST_XML_Converter\output\2025\processed\full_library\1_components_Inmaster_withsimsciid"),
#     Path(r"D:\NIST_XML_Converter\output\2025\processed\full_library\3_components_notInmaster_assignedsimsciid")
# ]

# OUTPUT_FILE = Path(r"D:\NIST_XML_Converter\prerequisites\Smiles_Prep_ICAS\1_TC_PC_Missing_Report.xlsx")


tc_missing = []
pc_missing = []


for folder in INPUT_FOLDERS:

    print(f"\nScanning: {folder}")

    for json_file in folder.rglob("*.json"):

        try:

            trcid = json_file.stem.split("_")[0]

            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            props = data.get("properties", {})

            tc = props.get("TC")
            pc = props.get("PC")

            if tc is None:
                tc_missing.append({
                    "TRCID": trcid,
                    "Property": "TC",
                    "Status": "PROPERTY_MISSING"
                })

            if pc is None:
                pc_missing.append({
                    "TRCID": trcid,
                    "Property": "PC",
                    "Status": "PROPERTY_MISSING"
                })

        except Exception as e:

            print(f"Error reading {json_file.name}: {e}")


tc_df = pd.DataFrame(tc_missing)
pc_df = pd.DataFrame(pc_missing)


if not tc_df.empty:
    tc_df.insert(0, "S.No", range(1, len(tc_df) + 1))

if not pc_df.empty:
    pc_df.insert(0, "S.No", range(1, len(pc_df) + 1))


with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    tc_df.to_excel(writer, sheet_name="TC_Missing", index=False)

    pc_df.to_excel(writer, sheet_name="PC_Missing", index=False)


print("\nDone")
print(f"TC Missing : {len(tc_df)}")
print(f"PC Missing : {len(pc_df)}")
print(f"Output     : {OUTPUT_FILE}")
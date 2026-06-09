"""
Script:
    37_generate_master_library_lb1_lb2_files.py

Purpose:
    This script generates a master NIST XML library file
    and builds final LB1/LB2 thermodynamic library files.

Functionality:
    - Reads processed XML component files
    - Sorts XML files using TRCID ordering
    - Generates master Bank-PURECOMP-NIST.xml library file
    - Adds XML component import entries
    - Formats and writes master XML library structure
    - Executes LMLibraryBuild utility
    - Generates final thermodynamic library files
    - Captures utility execution logs and errors

Input:
    Processed XML component files

Output:
    - Master Bank-PURECOMP-NIST.xml file
    - Generated LB1/LB2 thermodynamic library files
    - Utility execution logs
"""


from pathlib import Path
import os
import re
import subprocess
import time
from xml.etree.ElementTree import Element, SubElement, ElementTree


from config import (
    RUN_YEAR,
    PREREQ_DIR,
    OUTPUT_DIR,
    XML_LIBRARY_DIR,
    EXECUTABLES_DIR,
    LIBRARY_OUTPUT_DIR,
    ensure_directories
)

ensure_directories()

# ==================================================
# EXECUTABLE
# ==================================================

LM_LIBRARY_BUILD = EXECUTABLES_DIR / "LMLibraryBuild.exe"

# ==================================================
# OUTPUT DIRECTORIES
# ==================================================
#  Create final libraries output folder
LIBRARY_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ==================================================
# XML INPUT DIRECTORY
# ==================================================

XML_INPUT_DIR = XML_LIBRARY_DIR / "05_blacklist_processed"

# ==================================================
# VALIDATION 
# ==================================================

if not LM_LIBRARY_BUILD.exists():
    raise FileNotFoundError(f"Missing executable: {LM_LIBRARY_BUILD}")

if not XML_INPUT_DIR.exists():
    raise FileNotFoundError(f"Missing XML input folder: {XML_INPUT_DIR}")

# ==================================================
# DEBUG (optional)
# ==================================================

print("Executable:", LM_LIBRARY_BUILD)
print("XML input dir:", XML_INPUT_DIR)
print("Library output dir:", LIBRARY_OUTPUT_DIR)

# Create root <bank>
bank = Element("bank", attrib={"type": "PURECOMP", "name": "NIST"})

# Extract numeric part from file name (NSTxxxx)
def extract_number(filename):
    match = re.search(r'NST(\d+)', filename)
    return int(match.group(1)) if match else 0

# Get list of xml files
xml_files = [
    f for f in os.listdir(str(XML_INPUT_DIR))  
    if f.endswith(".xml") and f != "Bank-PURECOMP-NIST.xml"
]

#  Proper numeric sorting
xml_files.sort(key=extract_number)

# Add <import> elements
# for xml_file in xml_files:
#     SubElement(bank, "import", attrib={"file": xml_file})

import xml.etree.ElementTree as ET

# for xml_file in xml_files:
#     file_path = os.path.join(folder_path, xml_file)
    
#     tree = ET.parse(file_path)
#     comp_root = tree.getroot()   # <comp>

#     bank.append(comp_root)       # directly append

# Add <import> elements
for xml_file in xml_files:
    SubElement(bank, "import", attrib={"file": xml_file})

# Pretty print
def indent(elem, level=0):
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent(child, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

indent(bank)

# Save file
master_file_path = (
    XML_INPUT_DIR
    / "Bank-PURECOMP-NIST.xml"
)
ElementTree(bank).write(master_file_path, encoding='utf-8', xml_declaration=True)

print(f"Master XML created: {master_file_path}")

sleep_time = 5  # seconds
print(f"Waiting for {sleep_time} seconds before running LMLibraryBuild...")
time.sleep(sleep_time)

# Generating the lb1 & lb2 files.
# Command to run the utility tool with arguments
# command = [
#     r"D:\NIST_XML_Converter\psuedoinstall\FlashDriver\S4MThermo\Source\vc17\x64\Release\LMLibraryBuild.exe",
#     os.path.join(folder_path, "Bank-PURECOMP-NIST.xml"),
#     r"-lib=" + os.path.join(folder_path, "NISTL2025_SG60FdefaultExtrapol.lib"),
# ]
command = [

    str(
        LM_LIBRARY_BUILD
    ),

    str(
        XML_INPUT_DIR
        / "Bank-PURECOMP-NIST.xml"
    ),

    "-lib=" + str(
    LIBRARY_OUTPUT_DIR
    / f"NISTL{RUN_YEAR}_Full_Library.lib"
)

]

# command = [
#     r"D:\NIST_XML_Converter\psuedoinstall\FlashDriver\S4MThermo\Source\vc17\x64\Release\LMLibraryBuild.exe",
#     os.path.join(folder_path, "Bank-PURECOMP-NIST.xml"),
#     r"-lib=" + os.path.join(folder_path, "NISTL2025_Fillin_20Comp.lib"),r"-sse",
# ]

# Run the command and wait for it to finish
result = subprocess.run(command, capture_output=True, text=True)

# Print output and errors if any
print("STDOUT:")
print(result.stdout)

print("STDERR:")
print(result.stderr)

# Check if the process completed successfully
if result.returncode == 0:
    print("Conversion completed successfully.")
else:
    print(f"Conversion failed with return code {result.returncode}.")
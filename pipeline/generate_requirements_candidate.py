import ast
from pathlib import Path
import sys

# --------------------------------------------------
# Configuration
# --------------------------------------------------

PROJECT_PATH = Path(r"D:\NIST_XML_Converter\pipeline")
OUTPUT_FILE = PROJECT_PATH / "requirements_candidate.txt"

# --------------------------------------------------
# Standard Library Modules
# --------------------------------------------------

try:
    stdlib_modules = set(sys.stdlib_module_names)
except AttributeError:
    stdlib_modules = {
        "os", "sys", "json", "csv", "math", "re", "datetime",
        "pathlib", "collections", "itertools", "functools",
        "logging", "subprocess", "shutil", "typing",
        "tempfile", "glob", "argparse", "unittest",
        "xml", "sqlite3", "hashlib"
    }

# --------------------------------------------------
# Common Import → Package Mappings
# --------------------------------------------------

PACKAGE_MAPPING = {
    "dotenv": "python-dotenv",
    "PIL": "Pillow",
    "yaml": "PyYAML",
    "bs4": "beautifulsoup4",
    "cv2": "opencv-python",
    "sklearn": "scikit-learn",
}

# --------------------------------------------------
# Extract Imports
# --------------------------------------------------

packages = set()

for py_file in PROJECT_PATH.rglob("*.py"):
    try:
        with open(py_file, "r", encoding="utf-8") as f:
            tree = ast.parse(f.read(), filename=str(py_file))

        for node in ast.walk(tree):

            if isinstance(node, ast.Import):
                for alias in node.names:
                    pkg = alias.name.split(".")[0]
                    packages.add(pkg)

            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    pkg = node.module.split(".")[0]
                    packages.add(pkg)

    except Exception as e:
        print(f"Error reading {py_file}: {e}")

# --------------------------------------------------
# Filter Standard Library
# --------------------------------------------------

external_packages = set()

for pkg in packages:

    if pkg in stdlib_modules:
        continue

    external_packages.add(
        PACKAGE_MAPPING.get(pkg, pkg)
    )

# --------------------------------------------------
# Write Output
# --------------------------------------------------

external_packages = sorted(external_packages)

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    for pkg in external_packages:
        f.write(pkg + "\n")

# --------------------------------------------------
# Summary
# --------------------------------------------------

print(f"\nScanned folder: {PROJECT_PATH}")
print(f"Python files found: {len(list(PROJECT_PATH.rglob('*.py')))}")
print(f"External packages found: {len(external_packages)}")

print("\nPackages:")
for pkg in external_packages:
    print(f"  - {pkg}")

print(f"\nrequirements file generated:")
print(OUTPUT_FILE)
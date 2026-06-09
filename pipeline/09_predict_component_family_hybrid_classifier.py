"""
Script:
    09_predict_component_family_hybrid_classifier.py

Purpose:
    Predict component families using a hybrid classification
    approach based on names, formulas, and SMILES structures.

Functionality:
    - Reads component data with SMILES information
    - Classifies components using name-based rules
    - Classifies components using formula-based rules
    - Classifies components using RDKit structure analysis
    - Applies voting logic to assign final component groups
    - Preserves existing DIPPR family assignments when available
    - Generates predicted family classification output

Input:
    - Component data Excel file with SMILES information

Output:
    - Excel file with predicted component family classifications
"""

import json
import re
import pandas as pd
from collections import Counter
from openpyxl import load_workbook
from pathlib import Path

from rdkit import Chem
from rdkit.Chem import rdMolDescriptors
from rdkit.Chem.MolStandardize import rdMolStandardize

from rdkit import RDLogger
RDLogger.DisableLog("rdApp.error")

from utils import auto_adjust_column_widths

from config import (
    RUN_YEAR,
    PROCESSED_DIR,
    ensure_directories
    )

ensure_directories()

# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

INPUT_EXCEL = PROCESSED_DIR / "7_DIPPR_Family_Extraction_All_With_Smiles.xlsx"

OUTPUT_EXCEL = PROCESSED_DIR / "8_DIPPR_Family_Extraction_All_With_Smiles_none_Predicted_Family.xlsx"


# ==========================================================
# PRIORITY LIST (used for tie-breaking and RDKit classify)
# ==========================================================

PRIORITY = [
    # Strong functional groups
    "Peroxides", 
    "Anhydrides",
    "Acids",    
    "Amides",
    "Aldehydes", 
    "Esters",   
    "Ketones",
    "Silicon compounds",
    "Nitrogen compounds",
    
    
    "Amines",
    "Alcohols",

    # Strong element-based functional groups
    
    "Sulfur compounds",
    "Halogenated compounds",
    

    # Structural / heterocyclic groups
    "Aromatics",
    "Peroxides",  # above Aromatics
    "Ethers",        # above Napthenes
    
    
    # Hydrocarbon-only classes
    "Naphthenes",
    "Alkynes",
    "Olefins",
    "Paraffins",

    # Weakest classifications
    "Salts",
    "Elements",
    "Miscellaneous"
]


# ==========================================================
# Helper: element parsing & sets
# ==========================================================

ELEMENT_REGEX = re.compile(r"([A-Z][a-z]?)(\d*)")

METALS = {
    "Na", "K", "Ca", "Mg", "Al", "Fe", "Cu", "Zn", "Li",
    "Ba", "Sr", "Cr", "Mn", "Co", "Ni", "Ti", "V", "Sn"
}

HALOGENS = {"F", "Cl", "Br", "I"}


def parse_formula(formula: str) -> dict:
    """Parse a chemical formula into {element: count}."""
    if not isinstance(formula, str):
        return {}
    parts = ELEMENT_REGEX.findall(formula)
    counts = {}
    for elem, num in parts:
        counts[elem] = counts.get(elem, 0) + int(num or "1")
    return counts


def compute_hdi(counts: dict) -> float:
    """
    Compute Hydrogen Deficiency Index (HDI) from element counts.
    HDI = (2C + 2 + N - H - X)/2
    X = halogens (F,Cl,Br,I)
    """
    c = counts.get("C", 0)
    h = counts.get("H", 0)
    n = counts.get("N", 0)
    x = sum(counts.get(xe, 0) for xe in HALOGENS)
    if c == 0:
        return 0.0
    return (2 * c + 2 + n - h - x) / 2.0


# ==========================================================
# Enhanced NAME-based classification
# ==========================================================

def classify_by_name_enhanced(name: str) -> str:
    if not name:
        return "Unknown"

    n = name.strip().lower()
        # Peroxides / hydroperoxides
    if "peroxide" in n or "hydroperoxide" in n:
        return "Peroxides"

    # 0) ELEMENTS (simple exact matches)
    simple_elements = {
        "argon", "neon", "helium", "hydrogen", "oxygen",
        "nitrogen", "aluminum", "silicon", "sulfur"
    }
    if n in simple_elements:
        return "Elements"

    # 1) ACIDS
    if n.endswith("acid") or "oic acid" in n or "carboxylic acid" in n:
        return "Acids"

    # 2) ALCOHOLS
    # Alcohol inside parentheses or mid-name (generic)
    if "ol)" in n or "-ol" in n:
        return "Alcohols"
    if n.endswith("ol") and not n.endswith("thiol"):
        return "Alcohols"
    if "hydroxy" in n:
        return "Alcohols"
    if " alcohol" in n:
        return "Alcohols"

    # 3) ALDEHYDES
    if n.endswith("al") or "aldehyde" in n:
        return "Aldehydes"

    # 4) KETONES
# Ketone carbonyl must dominate unsaturation/aromatic hints
    if any(k in n for k in ["oxo", "acetyl", "benzoyl"]) or n.endswith("one"):
        return "Ketones"


    # SALTS (Generic rule: Metal + Anion Suffix)
    metal_tokens = [
        "sodium", "potassium", "calcium", "magnesium", "aluminum",
        "lithium", "zinc", "copper", "iron", "ferric", "ferrous",
        "nickel", "silver", "gold", "chromium", "manganese",
        "lead", "tin", "cobalt"
    ]
    anion_suffixes = ["ate", "ite", "ide"]
    if any(metal in n for metal in metal_tokens) and any(n.endswith(suf) for suf in anion_suffixes):
        return "Salts"

    # 5) SALT TOKENS (anion names)
    salt_tokens = [
        "sulfate", "sulphate",
        "sulfite", "sulphite",
        "nitrate", "nitrite",
        "chloride", "bromide", "fluoride", "iodide",
        "carbonate", "phosphate", "bicarbonate",
        "oxide", "hydroxide"
    ]
    if any(tok in n for tok in salt_tokens):
        return "Salts"
    # --------------------------------------------------------------
    # 6. ESTERS (careful, low‑impact rule)
    # --------------------------------------------------------------
    ester_anions = [
        "acetate", "ethanoate",
        "propanoate", "butanoate",
        "benzoate", "formate",
        "methanoate",
    ]

    has_ester_word = (
        " ester" in n or         # e.g. "butyl ethanoate ester"
        n.endswith(" ester") or  # name ends with " ester"
        "ester," in n            # "..., ester," in long names
    )

    has_carboxylate_like = any(tok in n for tok in ester_anions)

    if has_ester_word or has_carboxylate_like:
        return "Esters"
    
    # Aliphatic halogenated compounds dominate over alkene unsaturation
    if any(tok in n for tok in ["chloro", "bromo", "fluoro", "iodo"]):
        return "Halogenated compounds"
    # Isocyanates / diisocyanates → Nitrogen compounds
    if "isocyanate" in n:
        return "Nitrogen compounds"

    
    # Alkene keywords must dominate cycloalkane tokens
    # Alkene indicators (avoid alkane false positives like cycloheneicosane)
    # True alkenes (including cycloalkenes)
    # Alkene indicators including exocyclic methylene
    # if (
    #     n.endswith("ene")
    #     or any(k in n for k in ["enyl", "ylidene", "methylene", "ethenyl", "propenyl", "butenyl"])
    # ):
    #     return "Olefins"





    # 6) ESTERS (organic esters only)
    # ester_tokens = [
    #     "acetate", "propionate", "butyrate", "benzoate",
    #     "formate", "methacrylate", "acrylate"
    # ]
    # if "ester" in n or any(n.endswith(tok) for tok in ester_tokens):
    #     return "Esters"
    # if " ethanoate" in n or " methanoate" in n or " propanoate" in n:
    #     return "Esters"
    
    # 7) ETHERS
    # Alkoxy ethers (ethoxy, methoxy, tert-butoxy, etc.)
    alkoxy_tokens = [
        "methoxy", "ethoxy", "propoxy", "butoxy",
        "tert-butoxy", "isopropoxy", "alkoxy"
    ]

    if any(tok in n for tok in alkoxy_tokens):
        return "Ethers"
    # Generic alkoxy ethers: propenyloxy, pentoxy, hexoxy, etc.
    if re.search(r"\b[a-z0-9\-]+oxy\b", n):
        return "Ethers"


    if "ether" in n:
        return "Ethers"

   
    # # 2) Clear amine words → Amines
    # if any(k in n for k in ["amine", "aminium"]) and "amide" not in n:
    #     return "Amines"
    # # 9) AMIDES Clear amide words → Amides
    # if any(k in n for k in ["amide", "anilide", "imide", "urea", "carboxamide","carbamate", "carbanilate"]):
    #     return "Amides"

    # # 9) NITRO COMPOUNDS
    # if "nitro" in n:
    #     return "Nitrogen compounds"
    # # 3) Nitrogen‑compound patterns (not simple amines)
    # nitro_terms = ["nitro", "dinitro", "trinitro"]
    # nitrile_terms = ["nitrile", "cyanide", "cyano"]
    # ring_n_terms = [
    #     "pyridine", "pyridyl",
    #     "pyrimidine", "pyrimidyl",
    #     "pyrazine", "piperidine", "imidazole", "triazole",
    # ]
    # other_n_terms = ["azide", "azo", "diazo", "nitroso"]

    # if any(k in n for k in nitro_terms + nitrile_terms + ring_n_terms + other_n_terms):
    #     # avoid overriding explicit 'amine' names
    #     if "amine" not in n and "amino" not in n:
    #         return "Nitrogen compounds"
    #  # 8) AMINES
    # if "aniline" in n or n.endswith("amine") or "amino" in n:
    #     return "Amines"
    
    # # 1) Explicit amide names → Amides
    # if any(k in n for k in ["amide", "anilide", "formamide", "acetanilide", "benzamide"]):
    #     return "Amides"

    # # 2) Then nitrogen‑bias terms for Nitrogen compounds
    # nitrogen_bias_terms = [
    #     "nitrile", "cyanide", "cyano",
    #     "nitro", "nitrite", "nitrate",
    #     "azide", "azo", "diazo",
    #     "pyridine", "pyridyl", "picolinate", "isonicotinate",
    # ]
    # if any(k in n for k in nitrogen_bias_terms):
    #     if "amine" not in n and "amino" not in n:
    #         return "Nitrogen compounds"
    if "nitroaniline" in n:
        return "Amines"

# 1) Explicit amide / carbamate / urea names → Amides
    if any(k in n for k in [
        "amide", "anilide", "imide", "urea",
        "carboxamide", "carbamate", "carbanilate",
        "formamide", "acetanilide", "benzamide"
    ]):
        return "Amides"

    # 2) Nitro / nitrile / hetero N patterns → Nitrogen compounds
    nitrogen_bias_terms = [
        "nitro", "dinitro", "trinitro",
        "nitrile", "cyanide", "cyano",
        "nitrite", "nitrate",
        "azide", "azo", "diazo", "nitroso",
        "pyridine", "pyridyl", "picolinate", "isonicotinate",
        "pyrimidine", "pyrimidyl", "pyrazine",
        "piperidine", "imidazole", "triazole",
    ]

    if any(k in n for k in nitrogen_bias_terms):
        # do not override clear amine names
        if "amine" not in n and "amino" not in n:
            return "Nitrogen compounds"

    # 3) Clear amine names → Amines
    if any(k in n for k in ["amine", "aminium"]) and "amide" not in n:
        return "Amines"
    if "aniline" in n or n.endswith("amine") or "amino" in n:
        return "Amines"


    # 10) HALOGENATED COMPOUNDS
    if any(tok in n for tok in ["chloro", "bromo", "fluoro", "iodo"]):
        return "Halogenated compounds"

    # 11) AROMATICS
    arom_tokens = ["benz", "phenyl", "toluene", "xylene", "naphthalene"]
    if any(tok in n for tok in arom_tokens):
        return "Aromatics"
    
    # Fully hydrogenated aromatics → Napthenes
    if any(k in n for k in ["decahydro", "perhydro", "hexahydro", "octahydro"]):
        return "Naphthenes"


    # 12) NAPTHENES
    if "cyclo" in n and not any(tok in n for tok in ["benz", "phenyl", "naphth"]):
        return "Naphthenes"
    #Silicon Compounds
    if any(k in n for k in ["silane", "siloxane", "silyl", "disilane", "disiloxane"]):
        return "Silicon compounds"
    
    # ---- SULFIDES / THIOETHERS (HIGH PRIORITY) ----
    sulfide_tokens = [
        "thio",        # phenylthio, alkylthio
        "sulfide",
        "dithiane",
        "thiane",
        "thiol",
    ]   

    if any(tok in n for tok in sulfide_tokens):
        return "Sulfur compounds"
    


    return "Unknown"


# ==========================================================
# Enhanced FORMULA-based classification
# ==========================================================

def classify_by_formula_enhanced(formula: str) -> str:
    if not formula or not isinstance(formula, str):
        return "Unknown"

    counts = parse_formula(formula)
    if not counts:
        return "Unknown"

    elems = set(counts.keys())

    # 1) ELEMENTS
    if len(elems) == 1:
        return "Elements"

    # 2) SALTS (metal + non-metal)
    has_metal = any(m in elems for m in METALS)
    if has_metal and len(elems - METALS) > 0:
        return "Salts"

    # 3) SILICON COMPOUNDS
    if "Si" in elems:
        return "Silicon compounds"

    # 4) SULFUR COMPOUNDS (organic)
    if "S" in elems and "C" in elems:
        return "Sulfur compounds"

    # 5) HALOGENATED COMPOUNDS
    if any(h in formula for h in ["Cl", "Br", "F", "I"]):
        return "Halogenated compounds"

    # 6) ACIDS (COOH / CO2H)
    if "COOH" in formula or "CO2H" in formula:
        return "Acids"

    # 7) ALDEHYDES (CHO present, not acid)
    if "CHO" in formula and "COOH" not in formula and "CO2H" not in formula:
        return "Aldehydes"

    # 8) ESTERS (COO but not acid)
    if "COO" in formula and "COOH" not in formula and "CO2H" not in formula:
        return "Esters"

    # 9) PURE HYDROCARBONS – HDI logic
    if "C" in elems and "H" in elems and len(elems) == 2:
        hdi = compute_hdi(counts)
        if hdi <= 0.1:
            return "Paraffins"
        elif 0.9 <= hdi <= 1.1:
            return "Olefins"  # let RDKit handle olefin vs cyclo
        elif 1.9 <= hdi <= 2.1:
            # CnH2n-2 is ambiguous: alkynes OR dienes
            # Do NOT force Alkynes – let RDKit / Name decide
            return "Olefins"        
        elif hdi >= 4:
            return "Aromatics"

    return "Unknown"


# ==========================================================
# RDKit IN-PROCESS CLASSIFIER
# ==========================================================

# SMARTS patterns
RDKIT_PATTERNS = {
    "Peroxides": [
    "[OX2][OX2]",     # –O–O–
    "[OX2][OX2H]"     # –OOH (hydroperoxide)
],
    "Anhydrides": ["C(=O)OC(=O)"],
    # "Acids": ["C(=O)O[H]", "C(=O)[O-]"],
    # "Acids": ["C(=O)O", "C(=O)[O-]"],
    "Acids": [
    "[CX3](=O)[OX2H]",   # neutral COOH
    "[CX3](=O)[O-]"     # deprotonated
],
    "Esters": ["C(=O)O[#6]"],
    "Amides": ["C(=O)N", "C(=O)N([#6])"],
    "Aldehydes": ["[CX3H1](=O)[#6]"],
    "Ketones": ["C(=O)[#6;!$(C(=O)O)]"],
    "Alcohols": ["[OX2H]", "[OX2H2]"],
    "Ethers": ["[#6][OD2][#6]"],
    # "Amines": ["[NX3;H2,H1,H0;!$(NC=O)]"],
    "Amines": [
        # amines, but not amides / nitro / nitrile
        "[NX3;H2,H1,H0;!$(NC=O);!$([N+](=O)[O-]);!$([#6]#N)]",
    ],
        # Nitrogen compounds: all other N-types
    "Nitrogen compounds": [
        "[#6]#N",           # nitriles
        "[N+](=O)[O-]",     # nitro
        "N=N",              # azo
        "N#N",              # diazo/azide fragments (approx)
        "[#7;!$(NC=O)]",    # generic non‑amide nitrogen
    ],
    "Halogenated compounds": ["[F,Cl,Br,I]"],
    "Sulfur compounds": ["[#16]"],
    "Silicon compounds": ["[Si]"],
    "Aromatics": ["a"],
    "Olefins": ["C=C"],
    "Alkynes": ["C#C"], 
    "Paraffins": ["[CX4H0,CX4H1,CX4H2,CX4H3]"],
}

RDKIT_COMPILED = {
    g: [Chem.MolFromSmarts(s) for s in smarts]
    for g, smarts in RDKIT_PATTERNS.items()
}


def _is_aromatic(mol):
    for ring in mol.GetRingInfo().AtomRings():
        if all(mol.GetAtomWithIdx(i).GetIsAromatic() for i in ring):
            return True
    return False


def _is_naphthene(mol):
    if mol.GetRingInfo().NumRings() == 0:
        return False
    return not any(a.GetIsAromatic() for a in mol.GetAtoms())
# ==========================================================
# RDKit SALT / MIXTURE FIX
# ==========================================================

def normalize_smiles_remove_salts(smiles: str):
    """
    Normalize SMILES by removing salts / inorganic fragments
    and keeping the largest organic fragment.
    """
    if not smiles or not isinstance(smiles, str):
        return None

    try:
        mol = Chem.MolFromSmiles(smiles, sanitize=True)
        if mol is None:
            return None

        # Keep largest fragment (removes Na+, Cl-, etc.)
        chooser = rdMolStandardize.LargestFragmentChooser()
        mol = chooser.choose(mol)

        Chem.SanitizeMol(mol)
        return Chem.MolToSmiles(mol, canonical=True)

    except Exception:
        return None



def rdkit_classify(smiles: str) -> dict:
    if not smiles:
        return {"error": "Empty SMILES", "smiles": smiles}

    # mol = Chem.MolFromSmiles(smiles)
    # if mol is None:
    #     return {"error": "Invalid SMILES", "smiles": smiles}
    fixed_smiles = normalize_smiles_remove_salts(smiles)

    if not fixed_smiles:
        return {"error": "Invalid SMILES", "smiles": smiles}

    mol = Chem.MolFromSmiles(fixed_smiles)
    if mol is None:
        return {"error": "Invalid SMILES", "smiles": smiles}

    smiles = fixed_smiles


    found = []
    # ---- Lactone / cyclic ester override (HIGH CONFIDENCE) ----
    lactone = Chem.MolFromSmarts("[O;R][C](=O)[O;R]")
    if mol.HasSubstructMatch(lactone):
        return {
            "primary_group": "Esters",
            "all_groups": ["Esters"],
            "input_smiles": smiles,
            "smiles": smiles,
            "valid": True,
        }



    if _is_aromatic(mol):
        found.append("Aromatics")

    if _is_naphthene(mol):
        found.append("Naphthenes")

    for group, plist in RDKIT_COMPILED.items():
        for patt in plist:
            if patt and mol.HasSubstructMatch(patt):
                found.append(group)
                break

    mw = rdMolDescriptors.CalcExactMolWt(mol)
    # ---- Lactone / cyclic ester override (SAFE) ----
    if mol:
        for atom in mol.GetAtoms():
            if atom.GetAtomicNum() == 8:  # oxygen
                for bond in atom.GetBonds():
                    if bond.GetBondTypeAsDouble() == 2:  # C=O
                        # check if same O is in a ring
                        if atom.IsInRing():
                            found = ["Esters"]  # override to Ester
                            


        # Carboxylic acid must dominate ketone
    if "Acids" in found and "Ketones" in found:
        found.remove("Ketones")

    if not found:
        return {
            "primary_group": "Miscellaneous",
            "all_groups": ["Miscellaneous"],
            "input_smiles": smiles,
            "smiles": smiles,
            "mw": mw,
            "valid": True,
        }

    primary = next((p for p in PRIORITY if p in found), found[0])

    return {
        "primary_group": primary,
        "all_groups": list(dict.fromkeys(found)),
        "input_smiles": smiles,
        "smiles": smiles,
        "mw": mw,
        "valid": True,
    }


# ==========================================================
# VOTING ENGINE
# ==========================================================
def vote_groups(name_group, formula_group, rdkit_primary, rdkit_all_groups):
    groups = [name_group, formula_group, rdkit_primary]
    if isinstance(rdkit_all_groups, str):
        rdkit_groups = [g.strip() for g in rdkit_all_groups.split(",")]
    else:
        rdkit_groups = []


        # Special case: pure hydrocarbon tie → trust RDKit
    hydro_groups = ["Aromatics", "Naphthenes", "Olefins", "Paraffins"]

    if name_group in ["Unknown", "", None] and formula_group in hydro_groups and rdkit_primary in hydro_groups \
       and formula_group != rdkit_primary:
        return rdkit_primary
    # Peroxides dominate weak oxygen classes only
    # Peroxides dominate weak oxygen/structural classes
    if rdkit_primary == "Peroxides":
        if name_group in ["Alcohols", "Ethers", "Aromatics",
                          "Naphthenes", "Paraffins", "Salts",
                          "Unknown", "", None] \
           and formula_group not in ["Silicon compounds",
                                     "Sulfur compounds",
                                     "Nitrogen compounds"]:
            return "Peroxides"



    # Special case: Anhydrides → Acids
    if rdkit_primary == "Anhydrides":
        return "Acids"
    # Structural acid must override name-based ester
    if rdkit_primary == "Acids" and name_group == "Esters":
        return "Acids"   
# Aromatic core must dominate halogen substituents (RH family)
    if (
        rdkit_primary == "Halogenated compounds"
        and "Aromatics" in rdkit_groups
    ):
        return "Aromatics"
    
    # Ester dominance over acid (RS – Aromatic esters)
    if rdkit_primary == "Acids" and "Esters" in rdkit_groups:
        return "Esters"
    # Lactone / ester dominance over ketone
    if rdkit_primary == "Ketones" and "Esters" in rdkit_groups:
        return "Esters"
    # Structural ester overrides name-based alcohol
    if name_group == "Alcohols" and "Esters" in rdkit_groups:
        return "Esters"
    if "Anhydrides" in rdkit_groups and "Esters" in rdkit_groups:
        return "Esters"
    # Structural ester overrides name-based acid/alcohol noise
    if "Esters" in rdkit_groups and name_group in ["Acids", "Alcohols", "Ketones"]:
        return "Esters"

    
    # If RDKit and Name both say Alkynes, trust them over Formula
    if rdkit_primary == "Alkynes" and name_group == "Alkynes":
        return "Alkynes"

    # Optional: if RDKit says Alkynes and Formula says Olefins/Aromatics, still trust RDKit
    if rdkit_primary == "Alkynes" and formula_group in ["Olefins", "Aromatics"]:
        return "Alkynes"

        # --- Silicon bias: prefer Silicon compounds when clearly indicated ---
    if (formula_group == "Silicon compounds" or rdkit_primary == "Silicon compounds"):
        if name_group in ["Unknown", "Invalid", "", None, "Silicon compounds"]:
            return "Silicon compounds"
        # Silicon bias: if any signal says Silicon, prefer it over purely organic groups
    if "Silicon compounds" in [formula_group, rdkit_primary, name_group]:
        return "Silicon compounds"
    if rdkit_primary == "Invalid" and formula_group == "Silicon compounds":
    # If name is Aromatics or Halogenated but contains clear silicon keywords, keep Silicon
        return "Silicon compounds"
      
    # Sulfur functional group must dominate aromatic substituents
    if (
        rdkit_primary == "Sulfur compounds"
        and "Aromatics" in rdkit_groups
    ):
        return "Sulfur compounds"
    # Prevent neutral halides from being misclassified as salts
    if name_group == "Salts" and formula_group == "Halogenated compounds":
        return "Halogenated compounds"
    
    # Nitrogen dominance ONLY for derivatized carbonyl systems (imines / hydrazones)
    if (
        "Nitrogen compounds" in rdkit_groups
        and "Amines" not in rdkit_groups
    ):
        return "Nitrogen compounds"
    # Aromatic core must dominate when detected by name or formula
    # Unsaturated aliphatic systems must not be forced to Aromatics
    # Fully saturated cyclic systems must dominate over formula/name unsaturation
    # if rdkit_primary == "Napthenes":
    #     return "Napthenes"# Sulfur functional group must dominate aromatic substituents
 
    if (
        formula_group == "Aromatics"
        and name_group == "Olefins"
        and rdkit_primary == "Invalid"
    ):
        return "Olefins"
  

    # if (
    #     (name_group == "Aromatics" or formula_group == "Aromatics")
    #     and rdkit_primary == "Invalid"
    # ):
    #     return "Aromatics"
    # Aromatics fallback ONLY if name also says Aromatics
     # Fully saturated multiring hydrocarbons → Napthenes must dominate
    if (
        rdkit_primary == "Naphthenes"
        and formula_group in ["Olefins", "Aromatics", "Unknown"]
        and name_group in ["Aromatics", "Unknown"]
    ):
        return "Naphthenes"
    # Fully saturated multiring hydrocarbons → Napthenes must dominate
    if (
        name_group == "Naphthenes"
        and rdkit_primary == "Invalid"
        and formula_group in ["Olefins", "Aromatics", "Unknown"]
    ):
        return "Naphthenes"

    if (
        name_group == "Aromatics"
        and rdkit_primary == "Invalid"
    ):
        return "Aromatics"


    # Alkene functionality must dominate saturated ring (cycloalkanes)
    if "Olefins" in rdkit_groups and rdkit_primary == "Naphthenes":
        return "Olefins"
    # Multiring cycloalkanes: RDKit Napthenes must dominate
  








    # -------------------------------------------
    # STEP 1 — All Unknown → Miscellaneous
    # -------------------------------------------
    if all(g in ["Unknown", "Invalid", None, ""] for g in groups):
        return "Miscellaneous"

    # -------------------------------------------
    # STEP 2 — Remove Unknown for voting
    # -------------------------------------------
    valid_groups = [g for g in groups if g not in ["Unknown", "Invalid", None, ""]]

    # If only one valid → direct answer
    if len(valid_groups) == 1:
        return valid_groups[0]

    # Count votes
    votes = Counter(valid_groups)
    max_votes = max(votes.values())
    candidates = [g for g, c in votes.items() if c == max_votes]

    # -------------------------------------------
    # STEP 3 — Tie handling (minimally adjusted)
    # -------------------------------------------

    if len(candidates) > 1:
        # Treat these as strong functional groups from RDKit
        STRONG_RD_GROUPS = {
            "Acids", "Amides", "Aldehydes", "Ketones", "Esters",
            "Amines", "Alcohols", "Nitrogen compounds",
            "Sulfur compounds", "Halogenated compounds", "Silicon compounds",
        }

        HYDROCARBON_OR_STRUCT = {
            "Paraffins", "Olefins", "Naphthenes", "Aromatics",
            "Ethers", "Miscellaneous", "Elements", "Salts",
        }

        # 3.0 — If RDKit found a strong functional group and
        #        name & formula are only hydrocarbon/structural,
        #        let RDKit override.
        if (
            rdkit_primary in candidates
            and rdkit_primary in STRONG_RD_GROUPS
            and (name_group in HYDROCARBON_OR_STRUCT or name_group in ["Unknown", "", None, "Invalid"])
            and (formula_group in HYDROCARBON_OR_STRUCT or formula_group in ["Unknown", "", None, "Invalid"])
        ):
            return rdkit_primary

        # 3.1 — Original rule: Name wins if not unknown
        if name_group in candidates and name_group not in ["Unknown", "Invalid", None, ""]:
            return name_group

        # 3.2 — Else Formula wins if Name unknown
        if (name_group in ["Unknown", "Invalid", None, ""]) and \
           (formula_group in candidates and formula_group not in ["Unknown", "Invalid", None, ""]):
            return formula_group

        # 3.3 — Else RDKit wins if both name & formula unknown
        if (name_group in ["Unknown", "Invalid", None, ""]) and \
           (formula_group in ["Unknown", "Invalid", None, ""]) and \
           (rdkit_primary in candidates and rdkit_primary not in ["Unknown", "Invalid", None, ""]):
            return rdkit_primary

        # 3.4 — Else PRIORITY LIST
        for p in PRIORITY:
            if p in candidates:
                return p

        # 3.5 — Last fallback
        return candidates[0]

    # -------------------------------------------
    # STEP 4 — If not tie, return the only winner
    # -------------------------------------------
    return candidates[0]



# ==========================================================
# MAIN EXCEL PROCESSOR
# ==========================================================

def process_excel(input_path, output_path):
    df = pd.read_excel(input_path)

    df["SMILES"] = df["SMILES"].fillna("").astype(str).str.strip()

    name_groups = []
    formula_groups = []
    rdkit_primary_groups = []
    rdkit_all_groups = []
    voted_groups = []
    def is_missing(val):
        if val is None:
            return True
        if not isinstance(val, str):
            return False
        return val.strip() == "" or val.strip().lower() in {"none", "nan"}


    for _, row in df.iterrows():
        name = row.get("ComponentName", row.get("Component_Name", ""))
        formula = row.get("Formula", "")
        smiles = str(row.get("SMILES", "")).strip()

        # Name rule
        ng = classify_by_name_enhanced(name)
        name_groups.append(ng)

        # Formula rule
        fg = classify_by_formula_enhanced(formula)
        formula_groups.append(fg)

        # RDKit in-process classify
        rd = rdkit_classify(smiles) if smiles else {"primary_group": "Invalid", "all_groups": []}
        if "error" in rd:
            pg = "Invalid"
            ag = ""
        else:
            pg = rd.get("primary_group", "Invalid")
            ag = ", ".join(rd.get("all_groups", []))

        rdkit_primary_groups.append(pg)
        rdkit_all_groups.append(ag)

        # Voting
        dippr_family_flag = row.get("DIPPR_Family", "")
        final_main_group = row.get("Final_Main_Group", "")


        if is_missing(dippr_family_flag) or dippr_family_flag == "PO":
            voted = vote_groups(ng, fg, pg, ag)
        else:
            voted = final_main_group   # trust DIPPR, do NOT override

        voted_groups.append(voted)

    df["Name_Group"] = name_groups
    df["Formula_Group"] = formula_groups
    df["RDKit_Primary_Group"] = rdkit_primary_groups
    df["RDKit_Possible_Groups"] = rdkit_all_groups
    df["Hybrid_Voted_Group"] = voted_groups
    df["Group_Source"] = [
    "DIPPR" if not is_missing(row.get("DIPPR_Family")) else "HYBRID"
    for _, row in df.iterrows()]

    # def choose_main_group(row):
    #     fmg = row.get("Final_Main_Group")
    #     if isinstance(fmg, str) and fmg.strip() and fmg != "Miscellaneous":
    #         return fmg              # DIPPR wins
    #     return row["Hybrid_Voted_Group"]

    # df["Chosen_Main_Group"] = df.apply(choose_main_group, axis=1)

    df.to_excel(output_path, index=False)
    auto_adjust_column_widths(output_path)

    print(f"Hybrid voting classifier completed {output_path}")


# ==========================================================
# ENTRY POINT
# ==========================================================

if __name__ == "__main__":
    # process_excel(INPUT_EXCEL_Test_Master, OUTPUT_EXCEL_Test_Master)
    process_excel(INPUT_EXCEL, OUTPUT_EXCEL)

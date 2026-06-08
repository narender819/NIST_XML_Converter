import subprocess
import sys
import os
import time
from pathlib import Path


def run_script(script_name):
    print(f"Running {script_name}...")

    result = subprocess.run(
        [sys.executable, script_name],
        capture_output=True,
        text=True
    )

    print(result.stdout)

    if result.returncode != 0:
        print(f"Error running {script_name}:")
        print(result.stderr)
        return False

    return True


def main():

    scripts = [

        # ==================================================
        # LIBRARY EXTRACTION
        # ==================================================
        # "01_extract_library_xml_component_data.py",
        # "02_append_dippr_sponsor_sheet.py",

        # # ==================================================
        # # JSON / SMILES PROCESSING
        # # ==================================================
        # "03_extract_nist_json_and_smiles_data.py",
        # "04_clean_smiles_dataset.py",

        # # # ==================================================
        # # # DIPPR FAMILY PROCESSING
        # # # ==================================================
        # "05_extract_dippr_family_data.py",
        # "06_append_dippr_family_abbreviations.py",
        # "07_assign_dippr_family_groups.py",

        # # # ==================================================
        # # # COMPONENT CLASSIFICATION
        # # # ==================================================
        # "08_merge_smiles_with_component_data.py",
        # "09_predict_component_family_hybrid_classifier.py",

        # # # ==================================================
        # # # SIMSCI ID PROCESSING
        # # # ==================================================
        # "10_assign_simsci_ids.py",
        # "11_transform_nist_json_to_simsci_format.py",       
        # "12_update_processed_json_simsci_ids.py",

        # # ==================================================
        # # THERMODYNAMIC WORKFLOW
        # # ==================================================
        # "13_extract_nist_component_core_properties.py",
        # "14_extract_missing_fillin_properties.py",
        # "15_fill_missing_tc_pc_acentricfactor.py",
        # "16_build_nist_simsci_component_availability.py",
        "17_run_propeval_for_nist_components.py",
        "18_build_phasewise_pe1_enthalpy_workflow.py",
        "19_fill_missing_thermo_properties.py",
        "20_merge_nist_phasewise_with_simsci_metadata.py",
        "21_calculate_phasewise_trecon_temperatures.py",
        "22_cleanup_and_reorder_phasewise_sheets.py",
        # "23_generate_phasewise_trecon_propeval_runs.py",
        # "24_extract_trecon_propeval_results_to_excel.py",
        # "25_append_master_alias_names_to_json.py",

        # # ==================================================
        # # XML GENERATION
        # # ==================================================
        # "26_generate_simsci_xml_from_processed_json.py",

        # # ==================================================
        # # SCP PROCESSING
        # # ==================================================
        # "27_extract_nist_scp_correlation_coefficients.py",
        # "28_extract_simsci_scp_correlation_coefficients.py",
        # "29_calculate_scp_properties_from_correlations.py",

        # # ==================================================
        # # INTEGRATED CONSTANTS
        # # ==================================================
        # "30_calculate_scp_integrated_constants.py",
        # "31_calculate_lcp_integrated_constants.py",
        # "32_calculate_icp_integrated_constants.py",

        # # ==================================================
        # # XML FINALIZATION
        # # ==================================================
        # "33_update_xml_missing_property_refcodes.py",
        # "34_update_xml_enthalpy_integrated_constants.py",
        # "35_clear_duplicate_casnum_from_xml.py",
        # "36_process_blacklisted_component_xmls.py",

        # # ==================================================
        # # FINAL LIBRARY BUILD
        # # ==================================================
        # "37_generate_master_library_lb1_lb2_files.py",
    ]

    # ==================================================
    # PIPELINE DIRECTORY
    # ==================================================
    BASE_DIR = Path(r"D:\NIST_XML_Converter\pipeline")

    # ==================================================
    # EXECUTE PIPELINE
    # ==================================================
    for script in scripts:

        script_path = BASE_DIR / script

        success = run_script(str(script_path))

        if not success:
            print("Stopping pipeline due to error.")
            break

        else:
            print(f"Finished {script}. Waiting before next script...\n")

            time.sleep(4)

    else:
        print("All scripts executed successfully.")


if __name__ == "__main__":
    main()
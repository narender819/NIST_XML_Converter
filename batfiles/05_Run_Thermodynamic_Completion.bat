@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 15_fill_missing_tc_pc_acentricfactor.py || goto :error
python 16_build_nist_simsci_component_availability.py || goto :error
python 17_run_propeval_for_nist_components.py || goto :error
python 18_build_phasewise_pe1_enthalpy_workflow.py || goto :error
python 19_fill_missing_thermo_properties.py || goto :error
python 20_merge_nist_phasewise_with_simsci_metadata.py || goto :error
python 21_calculate_phasewise_trecon_temperatures.py || goto :error
python 22_cleanup_and_reorder_phasewise_sheets.py || goto :error
python 23_generate_phasewise_trecon_propeval_runs.py || goto :error
python 24_extract_trecon_propeval_results_to_excel.py || goto :error
python 25_append_master_alias_names_to_json.py || goto :error
python 26_generate_simsci_xml_from_processed_json.py || goto :error
python 27_extract_nist_scp_correlation_coefficients.py || goto :error
python 28_extract_simsci_scp_correlation_coefficients.py || goto :error
python 29_calculate_scp_properties_from_correlations.py || goto :error
python 30_calculate_scp_integrated_constants.py || goto :error
python 31_calculate_lcp_integrated_constants.py || goto :error
python 32_calculate_icp_integrated_constants.py || goto :error

echo.
echo Thermodynamic Completion Completed Successfully
pause
exit /b

:error
echo.
echo Thermodynamic Completion Failed
pause

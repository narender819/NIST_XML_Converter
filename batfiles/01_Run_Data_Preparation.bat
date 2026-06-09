@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 01_extract_library_xml_component_data.py || goto :error
python 02_append_dippr_sponsor_sheet.py || goto :error
python 03_extract_nist_json_and_smiles_data.py || goto :error
python 04_clean_smiles_dataset.py || goto :error
python 05_extract_dippr_family_data.py || goto :error
python 06_append_dippr_family_abbreviations.py || goto :error
python 07_assign_dippr_family_groups.py || goto :error
python 08_merge_smiles_with_component_data.py || goto :error
python 09_predict_component_family_hybrid_classifier.py || goto :error
python 10_assign_simsci_ids.py || goto :error
python 11_transform_nist_json_to_simsci_format.py || goto :error
python 12_update_processed_json_simsci_ids.py || goto :error
python 13_extract_nist_component_core_properties.py || goto :error
python 14_extract_missing_fillin_properties.py || goto :error

echo.
echo Data Preparation Completed Successfully
pause
exit /b

:error
echo.
echo Data Preparation Failed
pause
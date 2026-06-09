@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 26_generate_simsci_xml_from_processed_json.py || goto :error
python 33_update_xml_missing_property_refcodes.py || goto :error
python 34_update_xml_enthalpy_integrated_constants.py || goto :error
python 35_clear_duplicate_casnum_from_xml.py || goto :error
python 36_process_blacklisted_component_xmls.py || goto :error
python 37_generate_master_library_lb1_lb2_files.py || goto :error

echo.
echo Final Library Generation Completed Successfully
pause
exit /b

:error
echo.
echo Final Library Generation Failed
pause
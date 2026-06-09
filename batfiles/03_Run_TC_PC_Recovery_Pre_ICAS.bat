@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 39_generate_tc_pc_missing_report.py || goto :error
python 40_merge_tc_pc_missing_components.py || goto :error
python 41_append_smiles_to_tc_pc_missing.py || goto :error
python 42_generate_icas_smiles_batches.py || goto :error

echo.
echo ICAS Batch Files Generated Successfully
pause
exit /b

:error
echo.
echo TC_PC Recovery Failed
pause
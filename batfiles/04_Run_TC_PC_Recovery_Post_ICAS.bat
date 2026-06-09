@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 43_extract_tc_pc_omega_from_icas.py || goto :error

echo.
echo ICAS Results Imported Successfully
pause
exit /b

:error
echo.
echo ICAS Import Failed
pause
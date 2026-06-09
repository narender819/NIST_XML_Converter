@echo off

cd /d D:\NIST_XML_Converter\pipeline

python 26a_generate_intermediate_xml.py || goto :error
python 37a_generate_intermediate_library.py || goto :error

echo.
echo Intermediate Library Generation Completed
pause
exit /b

:error
echo.
echo Intermediate Library Generation Failed
pause
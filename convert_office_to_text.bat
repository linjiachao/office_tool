@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo Drag .xlsx or .docx files onto this bat to convert them.
echo If no files were dragged in, all .xlsx and .docx files in this folder will be converted.
echo.

python "%~dp0extract_office.py" %*
set EXIT_CODE=%ERRORLEVEL%

echo.
if "%EXIT_CODE%"=="0" (
    echo Conversion finished.
    echo Output folder: "%~dp0converted"
) else (
    echo Conversion finished with errors. Error code: %EXIT_CODE%
)
echo.
pause
exit /b %EXIT_CODE%

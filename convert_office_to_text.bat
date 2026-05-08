@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo Drag .xlsx/.docx files or a folder onto this bat to convert them.
echo If a folder is dragged in, Office files inside it will be converted into:
echo   sibling-folder\converted\folder-name\...
echo based on the dragged folder's parent directory, while keeping the original subfolder structure.
echo If no files or folders were dragged in, all .xlsx and .docx files in this folder will be converted.
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

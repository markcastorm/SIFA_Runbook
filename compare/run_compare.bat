@echo off
echo ============================================================
echo  SIFA Compare Tool
echo ============================================================
echo.
echo  Drop your reference file into reference_input/ then run me.
echo.

cd /d "%~dp0"

python compare.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo Comparison completed successfully.
) else (
    echo Comparison failed with error code %ERRORLEVEL%.
)

echo.
pause

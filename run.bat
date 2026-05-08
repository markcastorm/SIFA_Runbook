@echo off
echo ============================================================
echo  SIFA Pipeline - Fund Savings by Category (Quarterly)
echo ============================================================
echo.

cd /d "%~dp0"

python main.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo Pipeline completed successfully.
) else (
    echo Pipeline failed with error code %ERRORLEVEL%.
)

echo.
pause

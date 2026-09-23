@echo off
setlocal
cd /d "%~dp0"
where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw "test_clickable_invoice.pyw"
    exit /b 0
)

where python >nul 2>nul
if not errorlevel 1 (
    start "" python "test_clickable_invoice.py"
    exit /b 0
)

echo.
echo Python was not found on this computer.
echo Install Python 3 and try again.
pause
exit /b 1

@echo off
setlocal
cd /d "%~dp0"
where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw "..\clickable_supplier_launcher.py" "Southeren"
    exit /b 0
)
where python >nul 2>nul
if not errorlevel 1 (
    start "" python "..\clickable_supplier_launcher.py" "Southeren"
    exit /b 0
)

echo Python was not found. Please install Python 3 first.
pause
exit /b 1

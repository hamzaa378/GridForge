@echo off
setlocal
cd /d "%~dp0"
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
  echo pywin32 is not accessible to this Python interpreter.
  echo Open PowerShell, activate the environment where pywin32 is installed,
  echo then run: python scripts\excel_parity_506_desktop.py
  pause
  exit /b 2
)
python scripts\excel_parity_506_desktop.py
if errorlevel 1 (
  echo Excel test failed: see error above.
  pause
  exit /b 1
)
python scripts\excel_parity_506_compare.py
if errorlevel 1 (
  echo Comparison could not be completed. Please check the report error.
  pause
  exit /b 1
)
echo Done. See docs\PARITY_506_COMPARISON.md
pause

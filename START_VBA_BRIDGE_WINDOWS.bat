@echo off
setlocal
cd /d "%~dp0"
echo =====================================================
echo   GridForge OPTIONAL Excel VBA Bridge
 echo   CAUTION: Excel VBA can execute arbitrary code.
echo =====================================================
where py >nul 2>nul
if not errorlevel 1 (
 py -3 scripts\excel_vba_bridge.py
) else (
 python scripts\excel_vba_bridge.py
)
if errorlevel 1 (
 echo.
 echo Bridge not started. Requires Windows, Excel desktop,
 echo and: py -3 -m pip install pywin32
 pause
 exit /b 1
)
endlocal

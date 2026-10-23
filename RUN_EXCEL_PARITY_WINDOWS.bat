@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel% equ 0 (
  set "PYRUN=py -3"
) else (
  set "PYRUN=python"
)
%PYRUN% scripts\excel_parity_desktop.py
if errorlevel 1 (
 echo Excel execution failed or prerequisites are missing.
 echo Install desktop Excel and run: py -3 -m pip install pywin32
 pause
 exit /b 1
)
%PYRUN% scripts\excel_parity_compare.py
set EXITCODE=%errorlevel%
if %EXITCODE% equ 1 echo There are Excel vs GridForge mismatches. Read docs\PARITY_COMPARISON.md.
if %EXITCODE% equ 0 echo All comparable test cases matched on this Excel installation.
pause
exit /b %EXITCODE%

@echo off
setlocal
cd /d "%~dp0"
echo GridForge - Windows desktop installer builder
echo.
echo Prerequisites on this BUILDER: Python 3, Node.js with npm, Inno Setup 6.
echo The installed application itself requires none of those tools.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0desktop\build_windows.ps1"
set "result=%ERRORLEVEL%"
echo.
if "%result%"=="0" (
  echo SUCCESS. Find the installer in build\release\.
) else (
  echo BUILD FAILED with exit code %result%. See the output above.
)
pause
exit /b %result%

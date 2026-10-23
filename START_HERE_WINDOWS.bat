@echo off
setlocal
cd /d "%~dp0"
echo.
echo ====================================================
echo   GRIDFORGE - RUN WITH PYTHON ONLY

echo ====================================================
echo No npm, Node.js, pip packages, or database required.
echo.
where py >nul 2>nul
if not errorlevel 1 (
  py -3 run_gridforge.py
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python 3 was not found on PATH. See docs\SETUP_FROM_ZERO.md
    pause
    exit /b 1
  )
  python run_gridforge.py
)
if errorlevel 1 (
 echo.
 echo Launch failed. Review the error above.
 pause
 exit /b 1
)
endlocal

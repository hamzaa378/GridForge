@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
 echo Node.js 20 or newer is required. Install it from https://nodejs.org/
 pause
 exit /b 1
)
if not exist "dist\index.html" (
 echo Compiled application is missing. Run npm install followed by npm run build.
 pause
 exit /b 1
)
echo.
echo =========================================
echo   GridForge - GDG Frontend Task 2
 echo =========================================
echo.
echo Open in your browser: http://localhost:4173
echo Keep this window open while using GridForge.
echo Press Ctrl+C to stop.
echo.
node scripts\serve.mjs
pause

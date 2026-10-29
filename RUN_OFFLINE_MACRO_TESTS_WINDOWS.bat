@echo off
setlocal
cd /d "%~dp0"
echo GridForge offline macro engine - regression tests
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is required for developer tests, NOT to run the installed application.
  exit /b 1
)
call npm run build
if errorlevel 1 exit /b 1
node --test tests\macros_v6.test.mjs tests\macros_v3_offline.test.mjs
exit /b %errorlevel%

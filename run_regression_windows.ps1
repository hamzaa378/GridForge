# Run from PowerShell after extracting the repository.
# On first use: py -3 -m pip install playwright
#               py -3 -m playwright install chromium
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
npm run regression
if ($LASTEXITCODE -ne 0) { throw "GridForge regression failed. Review the first FAIL above." }
Write-Host 'GridForge full regression passed.' -ForegroundColor Green

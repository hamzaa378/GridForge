param(
    [switch]$SkipTests,
    [switch]$PortableOnly
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $root
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'Build on Windows x64, or use the Windows GitHub Actions workflow.'
}
Write-Host 'GridForge Windows desktop release build' -ForegroundColor Cyan
if (-not (Get-Command npm -ErrorAction SilentlyContinue)) { throw 'Node.js and npm required to BUILD. End users do not need them.' }
if (-not (Get-Command py -ErrorAction SilentlyContinue) -and -not (Get-Command python -ErrorAction SilentlyContinue)) { throw 'Python required to BUILD.' }
$venv = Join-Path $root '.desktop-build-venv'
if (-not (Test-Path (Join-Path $venv 'Scripts\python.exe'))) {
    if (Get-Command python -ErrorAction SilentlyContinue) { & python -m venv $venv }
    else { & py -3 -m venv $venv }
    if ($LASTEXITCODE -ne 0) { throw 'Unable to create build virtual environment' }
}
$python = Join-Path $venv 'Scripts\python.exe'
& $python -m pip install --disable-pip-version-check -r desktop/requirements-windows.txt
if ($LASTEXITCODE -ne 0) { throw 'Desktop dependency installation failed' }
& npm install --no-audit --no-fund
if ($LASTEXITCODE -ne 0) { throw 'npm install failed' }
if (-not $SkipTests) {
    & npm run regression:core
    if ($LASTEXITCODE -ne 0) { throw 'Regression tests failed; release not built' }
    & $python -m unittest discover -s tests -p 'test_desktop*.py' -v
    if ($LASTEXITCODE -ne 0) { throw 'Desktop unit tests failed; release not built' }
}
& npm run build
if ($LASTEXITCODE -ne 0) { throw 'TypeScript build failed' }
& $python desktop/gridforge_desktop.py --check
if ($LASTEXITCODE -ne 0) { throw 'Desktop assets missing' }
New-Item -ItemType Directory -Force -Path build/portable | Out-Null
$webAssets = (Resolve-Path (Join-Path $root 'dist')).Path
$portableDir = Join-Path $root 'build/portable'
$pyInstallerWork = Join-Path $root 'build/pyinstaller-work'
$pyInstallerSpec = Join-Path $root 'build/pyinstaller-spec'
if (-not (Test-Path (Join-Path $webAssets 'main.js')) -or -not (Test-Path (Join-Path $webAssets 'styles.css'))) { throw 'Compiled frontend assets are incomplete.' }
$appIcon = (Resolve-Path (Join-Path $root 'desktop/gridforge.ico')).Path
$desktopEntry = (Resolve-Path (Join-Path $root 'desktop/gridforge_desktop.py')).Path
$bridgeEntry = (Resolve-Path (Join-Path $root 'scripts/excel_vba_bridge.py')).Path
& $python -m PyInstaller --noconfirm --clean --onedir --windowed `
  --name GridForge --icon $appIcon `
  --distpath $portableDir --workpath $pyInstallerWork `
  --specpath $pyInstallerSpec --collect-all webview `
  --hidden-import webview.platforms.edgechromium `
  --hidden-import clr --add-data "$webAssets;dist" `
  $desktopEntry
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller desktop build failed' }
& $python -m PyInstaller --noconfirm --clean --onefile --console `
  --name GridForgeExcelBridge --icon $appIcon `
  --distpath $portableDir --workpath (Join-Path $root 'build/bridge-work') `
  --specpath $pyInstallerSpec `
  --hidden-import pythoncom --hidden-import win32timezone --hidden-import win32com.client `
  $bridgeEntry
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller VBA bridge build failed' }
if (-not (Test-Path 'build/portable/GridForge/GridForge.exe')) { throw 'GridForge.exe missing' }
if (-not (Test-Path 'build/portable/GridForgeExcelBridge.exe')) { throw 'Bridge EXE missing' }
if ($PortableOnly) {
    Write-Host 'PASS: portable folder at build/portable/GridForge' -ForegroundColor Green
    exit 0
}
$compiler = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 7\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $compiler) {
    $isccCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($isccCommand) { $compiler = $isccCommand.Source }
}
if (-not $compiler) {
    throw 'Inno Setup compiler not found. Install Inno Setup 6 (winget install JRSoftware.InnoSetup) or pass -PortableOnly.'
}
$releaseDir = Join-Path $root 'build/release'
New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
$buildDrive = $null
foreach ($letter in @('Z','Y','X','W','V','U','T','S','R','Q','P','O','N','M','L','K','J','I','H','G','F','E','D')) {
    if (-not (Get-PSDrive -Name $letter -ErrorAction SilentlyContinue)) {
        $buildDrive = $letter + ':'
        break
    }
}
if (-not $buildDrive) { throw 'No unused drive letter is available for the installer compiler.' }
& subst $buildDrive $root
if ($LASTEXITCODE -ne 0) { throw 'Unable to create a temporary short path for Inno Setup.' }
try {
    $installerScript = $buildDrive + '\desktop\installer\GridForge.iss'
    & $compiler $installerScript
    $compileCode = $LASTEXITCODE
} finally {
    & subst $buildDrive /D
}
if ($compileCode -ne 0) { throw 'Installer compilation failed' }
Write-Host 'PASS: Windows installer generated in build/release/' -ForegroundColor Green
Get-ChildItem build/release/*.exe | Select-Object Name, Length, FullName

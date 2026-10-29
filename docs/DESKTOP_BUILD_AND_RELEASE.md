# Build and ship GridForge Desktop (Windows x64)

## What is delivered

This repository includes **complete desktop wrapper source, build scripts and installer recipe**, but a real Windows executable can only be built and validated **on Windows**. The provided repository ZIP is **not** itself a `.exe` installer.

We use pywebview 6.2.1 (Windows Edge Chromium/WebView2 UI), PyInstaller 6.22.3 (bundles Python runtime, web frontend, and dependencies), and Inno Setup 6 (setup, Start menu, uninstall). The engine remains in TypeScript. We intentionally use PyInstaller **one-directory** mode for reliability and an Inno Setup EXE installer. Local workbooks are stored in the persistent WebView2 profile. The separate VBA bridge is packaged without changing Microsoft Excel Trust Center settings.

## Prerequisites (builder only)

On Windows 10/11 x64:
- Node.js 20+ with npm
- Python 3.12 x64 (recommended), `py` launcher, access to PyPI
- Inno Setup 6 (install with `winget install --id JRSoftware.InnoSetup -e`)
- Microsoft Edge WebView2 Runtime, for actual GUI smoke testing
- Optional: desktop Microsoft Excel for VBA bridge tests

## One-command local build

From the repository root in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\desktop\build_windows.ps1
```

The script creates an isolated `.desktop-build-venv`, installs pywebview/PyInstaller/pywin32 for packaging, installs npm dependencies, runs regressions, rebuilds TypeScript, freezes both executables, and compiles the Inno Setup installer.

For a code-only build that skips all regression and desktop unit-test phases, add `-SkipTests`:

```powershell
powershell -ExecutionPolicy Bypass -File .\desktop\build_windows.ps1 -SkipTests
```

Resulting files:

```text
  build/release/GridForge-Setup-11.26.0-win64.exe
build/portable/GridForge/GridForge.exe
build/portable/GridForgeExcelBridge.exe
```

If you don't have Inno Setup, use `powershell -ExecutionPolicy Bypass -File .\desktop\build_windows.ps1 -PortableOnly` to build portable executables first.

## Automated GitHub build (recommended)

1. Create an empty GitHub repository and upload/commit **the repository contents**, not the ZIP as one file.
2. Push the repository to GitHub (the source includes `.github/workflows/windows-desktop-release.yml`).
3. Go to **Actions > Windows desktop installer > Run workflow**.
4. After the run succeeds, open its **Artifacts** and download **GridForge-Windows-x64-Installer-and-Portable**.
5. The artifact contains the installer, portable ZIP, and `SHA256SUMS.txt`.
6. Test installation, launch, Excel import/export, light/dark theme, reopen/persistence, uninstall, and optional Excel VBA bridge on a real Windows x64 computer. After signing and validation, manually create a GitHub **Release** and attach the installer, portable ZIP and checksums for distribution; Actions artifacts expire.

The GitHub workflow performs an automated silent installer install/check/uninstall test and packages artifacts. This v11.26 build uses `-SkipTests`; no Windows runtime, installer launch, or live UI tests are run here. The workflow does **not** automatically publish to a public release, sign binaries, or claim interactive GUI tests passed. Distribution to third parties should follow your organization's signing/security process.

## Critical release acceptance checklist

- [ ] GitHub Actions Windows installer build succeeds.
- [ ] Installer opens cleanly on a fresh Windows 11 x64 VM.
- [ ] GridForge appears in Start menu; uninstall is available.
- [ ] Embedded UI loads with no white screen/JavaScript errors.
- [ ] Light/dark theme selection persists across app relaunch.
- [ ] Editing cells and `=SUM(1,2)` works.
- [ ] XLSX open -> edit -> native Save As -> reopen works.
- [ ] XLSM open -> edit -> Save XLSM preserves original VBA project (without running it).
- [ ] Multiple sheets, undo/redo, right-click, formula browser and ribbon still work.
- [ ] Workbook survives app close/reopen; JSON backup round-trip works.
- [ ] Optional VBA bridge connects only when explicitly run; trusted macro works with installed desktop Excel.
- [ ] Uninstall removes program files and Start menu entry while intentionally retaining user profile.
- [ ] Installer SHA-256 recorded. Malware/SmartScreen reviews and signing handled before broad release.

## Known limitations

- Only Windows x64 is targeted by this packaging recipe.
- Microsoft Edge WebView2 runtime is an OS prerequisite, not bundled.
- Native file picker and installer appearance need acceptance testing on real Windows. The source-only/Linux checks cannot verify it.
- The workbook engine remains a subset of Excel; no full Excel parity claim.
- The optional VBA bridge uses desktop Microsoft Excel; VBA is not run by the embedded browser engine.
- App data is stored per Windows user; the portable build and installer share the same profile when launched under that user.
- This is an **unsigned** installer until you supply a trusted signing certificate.

## v10.6 Macro Studio

The installer also adds a second Start menu shortcut: **GridForge Macro Studio**. It launches the same `GridForge.exe` using `--macro-workbench` and opens the offline macro dialog. No separate VBA or Node runtime is required by end users.

For the release checklist, verify: selecting the macro-specific shortcut, importing `examples/macros/parameterized_tax.json` after loading an appropriate workbook, previewing edits with `{"tax":"0.18","minimum":"500"}`, running, undoing, exporting the JSON macro, and closing/reopening the app.

The `.exe` and installer must be compiled on Windows; this repository ZIP is a build kit, not an installer binary.

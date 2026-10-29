# GridForge 11.4 Desktop for Windows

GridForge Desktop packages the existing v10.4 spreadsheet frontend as a Windows application. Formula evaluation, ribbon tabs, themes, sheet operations and XLSX/XLSM support are unchanged.

## What the user installs

- **GridForge-Setup-10.4.0-win64.exe**: ordinary Windows installer with Start menu shortcut, optional desktop shortcut, and uninstall entry
- **GridForge-Portable-win64.zip**: alternative standalone folder for non-installer use. Extract **the entire folder**, then run `GridForge.exe`
- **GridForgeExcelBridge.exe**: separately run optional bridge for **trusted VBA macros only**. Excel desktop must be installed and able to run the macro under its Trust Center settings.

Neither package requires end users to install Node.js, npm, Python or pip packages. **Microsoft Edge WebView2 Runtime** is required on Windows (commonly present on Windows 11). If missing, install Microsoft's official runtime; it is not bundled in this release.

## Install and run

1. Use Windows 10 or 11, 64-bit, with Microsoft Edge WebView2 Runtime installed.
2. Open `GridForge-Setup-10.4.0-win64.exe` and follow the installer.
3. Choose whether to create a desktop shortcut.
4. Launch **GridForge** from the Start menu.
5. Open a supported `.xlsx`, `.xlsm`, `.csv` or GridForge `.json` using the **File > Open** command (`Ctrl+O`).
6. Save to disk with **Ctrl+S** for XLSX, or use File menu exports. The desktop app opens a **native Windows Save As dialog**.

The app starts a loopback-only local server internally on an available port at `127.0.0.1`. This is normal and does **not** mean it needs Internet access. Each launch picks its own port, so simultaneous or stale instances do not contend for one fixed port.

## Autosave and data retention

Edits are stored in the Windows embedded browser profile at `%LOCALAPPDATA%\GridForge\WebView2`, with a stable local origin. The profile is **different from GridForge opened in Chrome/Edge**; existing browser-local workbooks will not automatically migrate. Export them from the browser version to `.json` or `.xlsx`, then open in the desktop version.

Always export important workbooks to files: embedded browser storage is not a replacement for explicit backups. The uninstaller intentionally leaves the profile and workbooks on disk. To erase it manually, back up important files, close GridForge, then remove `%LOCALAPPDATA%\GridForge`.

## Important compatibility notes

GridForge has **507 locally evaluated or workbook-aware registered formulas** plus **9 connected-function names** that require services and are not fully operational offline. The tested Excel parity does not establish full Excel equivalence. XLSM preserves VBA where supported; VBA **does not execute inside GridForge itself**.

To use Excel VBA: start the separate `GridForgeExcelBridge.exe`, copy its random token to the GridForge **Automate** panel, select an existing trusted VBA procedure and explicitly approve execution. This requires Microsoft Excel desktop and is not automatic. Never run untrusted macros.

## Troubleshooting

- Blank window or WebView error: confirm Microsoft Edge WebView2 Runtime is installed and current. Restart GridForge.
- Failed startup: confirm Microsoft Edge WebView2 Runtime is installed, then restart GridForge. The app chooses a free local port automatically.
- Old browser workbook not showing: this is a separate desktop profile. Export from the browser copy and import to desktop.
- SmartScreen shows an unknown publisher: the example release is **unsigned**. For public distribution, sign the installer and executables with a trusted code-signing certificate. Don't tell users to bypass organizational policies.
- Macros unavailable: install desktop Excel and use the separately launched, user-approved Excel VBA bridge. WebView2 alone does not execute VBA.
- Incorrect formulas or unsupported Excel workbook features: refer to the bundled compatibility documentation; report a minimal reproduction.

## Uninstall

Windows Settings > Apps > Installed apps > GridForge > Uninstall, or use **Uninstall GridForge** in the Start menu. Workbook data in `%LOCALAPPDATA%\GridForge` remains unless manually removed.

## Macro Studio (offline)

Open **Start → GridForge → GridForge Macro Studio** or choose Automate in the main GridForge window. Import a `.json` macro, enter any parameter overrides, click Preview, then Run. This executes the GridForge v4 declarative action engine, not Microsoft VBA, and needs no Excel installation. See `MACRO_STUDIO_V4.md`.

## v10.7 charts and visual formatting
Select data on a worksheet, open **Insert**, and select Column, Bar, Line, Area, Pie, or Scatter. Edit the title/source in the preview dialog, then drag the chart header to move it or its bottom-right corner to resize. Charts are stored in GridForge JSON workbooks and can be downloaded as SVG images. Editable charts are not currently exported to XLSX or XLSM as native Excel chart objects; export JSON to preserve them. The Home tab offers font, size, color, cell borders, wrap text, strikethrough, and cell-style presets.

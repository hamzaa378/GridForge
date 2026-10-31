# GridForge 12 Desktop for Windows

GridForge Desktop packages the GridForge spreadsheet application as a Windows program. It includes an Excel-inspired ribbon, formula engine, charts, PivotTables, workbook tools, and XLSX/XLSM support. Excel compatibility is broad but incomplete; see the compatibility audit for known limits.

## What the user installs

- **GridForge-Setup-<version>-win64.exe**: ordinary Windows installer with Start menu shortcut, optional desktop shortcut, and uninstall entry
- **GridForge-Portable-win64.zip**: alternative standalone folder for non-installer use. Extract **the entire folder**, then run `GridForge.exe`
- **GridForgeExcelBridge.exe**: separately run optional bridge for **trusted VBA macros only**. Excel desktop must be installed and able to run the macro under its Trust Center settings.

Neither package requires end users to install Node.js, npm, Python or pip packages. **Microsoft Edge WebView2 Runtime** is required on Windows (commonly present on Windows 11). If missing, install Microsoft's official runtime; it is not bundled in this release.

## Install and run

1. Use Windows 10 or 11, 64-bit, with Microsoft Edge WebView2 Runtime installed.
2. Open the versioned `GridForge-Setup-<version>-win64.exe` installer and follow the prompts.
3. Choose whether to create a desktop shortcut.
4. Launch **GridForge** from the Start menu.
5. Open a supported `.xlsx`, `.xlsm`, `.csv` or GridForge `.json` using the **File > Open** command (`Ctrl+O`).
6. Save to disk with **Ctrl+S** for XLSX, or use File menu exports. The desktop app opens a **native Windows Save As dialog**.

The app starts a loopback-only local server internally on an available port at `127.0.0.1`. This is normal and does **not** mean it needs Internet access. Each launch picks its own port, so simultaneous or stale instances do not contend for one fixed port.

## Autosave and data retention

Edits are stored in the Windows embedded browser profile at `%LOCALAPPDATA%\GridForge\WebView2`, with a stable local origin. The profile is **different from GridForge opened in Chrome/Edge**; existing browser-local workbooks will not automatically migrate. Export them from the browser version to `.json` or `.xlsx`, then open in the desktop version.

Always export important workbooks to files: embedded browser storage is not a replacement for explicit backups. The uninstaller intentionally leaves the profile and workbooks on disk. To erase it manually, back up important files, close GridForge, then remove `%LOCALAPPDATA%\GridForge`.

## Defined names

Use **Formulas > Define Name** to create a name for the active worksheet or the entire workbook. Names can refer to a cell or range, or use a supported formula such as `=SUM(Sheet1!A1:A10)` or a `LAMBDA`. Worksheet-scoped names take precedence over workbook-scoped names with the same spelling. The Name Manager displays comments and hidden-name metadata. References inside both scopes update when worksheets are renamed, and become `#REF!` if the referenced sheet is deleted. Excel's complete defined-name semantics and external workbook links are not supported.

## Worksheet capacity

GridForge keeps worksheets sparse and supports a canvas up to 10,000 rows by 256 columns. The default blank canvas is smaller (400 rows by 52 columns). Each worksheet is capped at 500,000 populated or formatted cells, XLSX import at 250,000 cell records per sheet, and clipboard actions at 50,000 cells per operation. Clearing selected content or formatting uses sparse entries to avoid creating work for every blank cell. Some tools have lower limits. This is larger than earlier GridForge releases but does not match Excel's maximum worksheet size.

## Important compatibility notes

GridForge has **506 locally evaluated or workbook-aware registered formulas** plus **10 connected-function names** that require services and are not fully operational offline. Dynamic arrays spill from supported formula results; unary and binary operators work element-wise on matching bounded array shapes. The per-formula output limit is 20,000 cells, and some Excel array error/shape semantics differ. The tested Excel parity does not establish full Excel equivalence. XLSM preserves VBA where supported; VBA **does not execute inside GridForge itself**.

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
Select data on a worksheet, open **Insert**, and select Column, Bar, Line, Area, Pie, or Scatter. Edit the title/source in the preview dialog, then drag the chart header to move it or its bottom-right corner to resize. Charts can be downloaded as SVG or high-resolution 2x PNG images. Supported chart types export as editable native Excel chart objects; specialized ChartEx layouts include Histogram, Waterfall, Pareto, Box & Whisker, Funnel, Treemap, and Sunburst. Box & Whisker charts offer inclusive/exclusive quartiles, inner/outlier points, and mean markers/lines. For supported category charts, **Chart Design > Data Table** shows source values below the plot; **Table Keys** adds each series color and name. The SVG preview abbreviates very large tables, while XLSX export writes the native chart data-table option. **Plot Hidden Data** restores filtered and manually hidden worksheet rows or columns to supported standard charts; it round-trips through Excel chart settings. Advanced chart import and editing are still partial; see the compatibility audit. Use **Accessibility description** in the chart editor to add concise alt text for screen readers; it is kept in GridForge workbooks and supported Excel chart drawing metadata.The Home tab offers font, size, color, cell borders, wrap text, strikethrough, and cell-style presets.

## Freeze worksheet panes

On the **View** tab, select the cell below and to the right of the rows and columns you want to keep visible, then choose **Freeze Panes**. Choose **Unfreeze Panes** to clear the frozen rows and columns. The existing **Freeze row** and **Freeze col** buttons still toggle the top row and first column.

# GridForge v11.26 — Standalone Excel-style Spreadsheet Workspace

GridForge is an offline-first spreadsheet application with its own worksheet grid, formula engine and Excel-inspired ribbon; Microsoft Excel is not required. Version 11.26 improves Excel-compatible date calculations, including U.S. and European 30/360 methods, Actual/Actual year fractions, and Excel date-range checks.

## v11.26 highlights
- `DAYS360` now distinguishes Excel's U.S. (NASD) and European methods, including month-end adjustments. `YEARFRAC` basis 0 and 4 use their matching conventions; basis 1 accumulates actual days against each year's serial-day length.
- Date serial validation now follows Excel's supported 1900-system range through 9999-12-31 in shared date helpers and working-day functions, including `WORKDAY.INTL` / `NETWORKDAYS.INTL`, while retaining serial 60 for Excel's fictitious 1900-02-29. Invalid all-weekend masks are rejected by `WORKDAY.INTL`.
- Formula references are stronger: `INDIRECT` handles A1/R1C1 references, local names and cross-sheet ranges; `OFFSET` now returns tracked rectangular references; `ADDRESS` supports A1/R1C1 and relative/absolute modes.
- `CELL` and `ISFORMULA`/`FORMULATEXT` inspect supported references across worksheets, while filter-aware `SUBTOTAL`/`AGGREGATE` refresh when row visibility changes.
- Scatter charts now plot multiple Y series against numeric X values and support horizontal axis bounds, units and formats that round-trip in supported XLSX charts.
- Pie and doughnut charts add first-slice rotation, hole sizing, per-slice colors, label placement and supported XLSX round-tripping.

- Data tab includes clear entry points for Power Query, the Data Model, Power Pivot/DAX, OLAP, native Timelines, and PivotChart field buttons. Selecting one opens a notice that explains what is missing and what GridForge supports instead.
- Help → Feature availability lists major compatibility gaps, partial implementations, and alternatives, including formula compatibility, specialized chart interchange, PivotCharts, and external refresh.
- Pivot Summary field counts follow source dimensions and available worksheet/output space instead of the previous fixed eight-field cap. The source picker accepts the engine's full 104,000-cell source bound. Each selected row or column field can now use its own Values calculation for ascending or descending value sorting, and those sort settings round-trip through native PivotTable AutoSort metadata.
- Native XLSX PivotTable import, summary calculation, and export support a separate Top/Bottom value filter on each selected row or column field. The Pivot Summary dialog can add, edit, and remove multiple filters, each with its own field, Values calculation, direction, and item count.
- Pivot item filters distinguish no filter from an explicit empty selection, preserving imported Excel states where no items are visible. Slicer items toggle independently while multi-selecting, and clearing a slicer restores all items.
- XLSX PivotTable import resolves workbook-scoped names that point directly to worksheet cell ranges, in addition to supported tables and local names. Formula-based and external-workbook names remain unsupported.
- Compatibility follow-up: supported Excel date/time number formats now round-trip through XLSX, and 1904-system date literals normalize on import. Histogram bin width, underflow, and overflow settings round-trip through native ChartEx; ChartEx legends preserve placement; Waterfall can round-trip a final total point. Formula adjustments cover XLOOKUP row/column arrays, selected date/time and blank/count behavior, NUMBERVALUE separators/percent signs, ASCII-space TRIM, DATEDIF boundaries, and XNPV validation; these remain targeted fixes rather than full Excel equivalence.
- Lookup compatibility now includes `MATCH` exact and sorted approximate modes, `XMATCH` and `XLOOKUP` forward/reverse and sorted binary searches, plus XLOOKUP wildcard and next-smaller/next-larger modes. Binary search validates ordering; broad Excel comparison across coercion and locale edge cases remains outstanding.
- Waterfall charts accept multiple total/subtotal point indexes in their editor and persist supported indexes using native ChartEx subtotal metadata. Page Setup now imports/exports Print Titles; repeated rows also repeat in browser print, while repeated columns are retained for Excel.
- Combo chart XLSX import/export now retains mixed column/line series and their primary/secondary axis assignments. Validation rules now preserve Excel input prompts and stop/warning/information alerts, with matching controls in the editor.
- `TEXTSPLIT` accepts delimiter arrays, case-insensitive matching, ignore-empty handling for row/column delimiters, and custom padding for uneven rows. Excel’s default error padding and some locale-sensitive casing remain incomplete.
- VLOOKUP/HLOOKUP and LOOKUP approximate modes now validate ascending order; TEXTBEFORE/TEXTAFTER support directional instances, case mode, match-at-end, fallback, and empty-input rules. View → Gridlines now controls persisted worksheet gridline visibility while keeping formatted cell borders.
- Standard chart label positions are editable and round-trip through supported DrawingML, and series names matching headers export as linked cell references. Best Fit rendering is approximate; ChartEx label positions remain unsupported.
- The cell context menu now inserts, edits, or removes worksheet hyperlinks. Imported, created, and direct `HYPERLINK` formula targets render with hyperlink styling and open supported external targets or select in-workbook cell ranges on visible worksheets. Nested formula links and named destinations remain unsupported.
- Standard chart labels can show values, categories, series names, and pie/doughnut percentages. `SHEET()` and `SHEETS()` use workbook worksheet order and count for supported single-sheet references; selected `AGGREGATE` statistical operations now include median, mode, and percentile/quartile variants.
- Cell formatting now includes vertical top/middle alignment and indent controls, with supported settings preserved in XLSX.
- Custom Excel number-format codes can be entered in the formatting toolbar and preserved through XLSX round-trips; the grid renders a supported subset and safely falls back for unsupported tokens.
- Conditional formatting now supports optional three-color scales with a 50th-percentile midpoint, preserved in native XLSX color-scale rules.
- Conditional numeric rules include inclusive comparisons, not-equal, and between/not-between, with supported operators preserved in XLSX.
- Dynamic-array `FILTER` can include rows or columns and return an array for an empty result; `SORT` supports sorting across columns, and `SORTBY` supports multiple stable keys with independent directions.
- Dynamic-array `TAKE` and `DROP` distinguish omitted arguments from explicit zero counts and return `#CALC!` for an empty result.
- Imported PivotTables retain manually ordered field items during refresh and XLSX export. Choose “Manual” in the Pivot Summary dialog to reorder items with Move Up/Move Down controls; manually grouped fields list their group captions, and date or numeric grouping lists the displayed group labels instead of raw source values. Opening an existing Pivot after source data changes, or changing its source range, reconciles added and removed categories while retaining the relative order of existing items; new values follow source order. “Manual (starts in source order)” retains the original encounter order until you reorder items.
- Source release includes the TypeScript source, compiled browser app (`dist/`), desktop wrapper, installer recipe, GitHub Actions workflow, and complete dependency/build instructions.
- Compiled release artifacts: `build/release/GridForge-Setup-11.26.0-win64.exe` (installer), `GridForge-Portable-win64.zip` (portable app), and `GridForge-v11.26.0-source.zip` (GitHub-ready source with compiled `dist/`). SHA-256 values are in `build/release/SHA256SUMS.txt`.

See [v11.26 feature matrix](docs/FEATURE_VERIFICATION_V11_26.md), [release notes](docs/RELEASE_NOTES_V11_26.md), and [desktop build guide](docs/DESKTOP_BUILD_AND_RELEASE.md).

## v11.22 highlights

- Define manual PivotTable groups with `Group = item | item`, one group per line; unassigned values remain visible as independent categories.
- Manual groups feed the summary, filters and native XLSX PivotTable discrete grouping metadata. Supported imports restore group captions and item mappings.
- TypeScript check/build and installer compilation only; no automated tests, Windows runtime checks, live UI checks or Excel GUI validation.
- See the [v11.22 feature matrix](docs/FEATURE_VERIFICATION_V11_22.md) and [release notes](docs/RELEASE_NOTES_V11_22.md).
## v11.15 highlights

- Create stock charts from Open, High, Low and Close columns in any order, with an optional date/label column.
- Invalid high/low bounds show a chart message, and incomplete rows are skipped.
- The Inno Setup installer and desktop metadata are updated to version 11.15.0; the installer was compiled without running Windows test phases.
- Stock chart XLSX export currently uses a clustered-column fallback.
- See the [v11.15 feature matrix](docs/FEATURE_VERIFICATION_V11_15.md) and [release notes](docs/RELEASE_NOTES_V11_15.md).

## v11.14 highlights

- Add **Funnel** charts for sequential stages and **Treemap** charts using one or more category columns plus a numeric value column.
- Create a **Descriptive Statistics** report with mean, standard error, median, mode, sample variance and standard deviation, kurtosis, skewness, range, min, max, sum and count.
- Statistics are written to a separate output range; text, blanks and formula errors are ignored, and overlapping the source is rejected.
- Specialized charts still export to XLSX as clustered columns. See the [v11.14 feature matrix](docs/FEATURE_VERIFICATION_V11_14.md) and [release notes](docs/RELEASE_NOTES_V11_14.md).

## v11.13 highlights

- Create Pareto charts with categories ranked by contribution, an accumulated percentage line, and a combined “Other” category for long tails.
- Create box-and-whisker charts with interpolated quartiles, Tukey whiskers and outlier markers.
- Set histogram bin count from 1 to 50, or leave it at 0 for automatic bins.
- Use Goal Seek to preview a numeric input that drives a formula toward a target. Calculations run on workbook copies; apply checks that the workbook has not changed since preview.
- Histogram now exports and imports as native ChartEx with a source range and either bin count or bin width for one numeric series, plus title, data-label, and series-color settings. Treemap and Sunburst preserve multi-level category columns; other advanced hierarchy/style behavior remains limited.
- Refresh public HTTPS CSV and JSON query sources from the Data tab. Endpoints must allow browser CORS; requests send no credentials, reject redirects, and are capped at 3 MB. XLSX exports include the last fetched result as cell values, not as an Excel connection.
- Goal Seek is bounded to one changing cell on the active worksheet and a finite numeric target; it is not a general Excel Solver replacement.

See the [v11.13 feature matrix](docs/FEATURE_VERIFICATION_V11_13.md), [release notes](docs/RELEASE_NOTES_V11_13.md) and [continuation guide](docs/CONTINUATION_GUIDE_V11_13.md).

## v11.12 highlights

- Use direct references across worksheets, including quoted names and ranges, such as `=Sheet2!A1` or `=SUM('Quarter 1'!B2:B20)`.
- Create Pivot Summary reports, including Values-only summaries with no row fields, with up to eight nested fields in each of Rows and Columns with independently reorderable Rows, Columns, and primary Values fields, up to eight value calculations (including repeated source fields and custom captions), independent Sum, Count, Count Numbers, Average, Minimum, Maximum, Product, StdDev, StdDevp, Var and VarP aggregations, with Excel-like numeric-versus-text defaults, Show Values As percentages of grand, row, column, parent-row, parent-column, % Of the Previous or Next item, running-total, percent-running-total, Difference From, % Difference From, and rank ascending/descending within a selected Base field, and Index, hierarchical row subtotals, tabular, outline, or compact report layout, subtotals above or below row groups, with optional blank lines after parent groups, manual categorical grouping for selected row/column items, plus one-level date grouping by year, quarter, month, or day and numeric interval grouping with data-derived or explicitly set bounds, for a selected row/column field, a global sort default with independent ascending/descending/source-order settings for each row and column field, value sorting by a selected calculation, numeric/date-aware field sorting, and Top/Bottom item filters ranked within each parent group by a selected Values calculation, optional row/column grand totals, row/column item filters and caption filters (including inclusive Between/Not between), and up to eight searchable multi-item report filters. Inclusive date-range filters can be applied to ungrouped row or column fields and are stored with the Pivot Summary; XLSX export represents them as PivotField item visibility. Large item lists are searched in bounded chunks. Blank row and column items are retained, and ordinary edits refresh only summaries whose sources changed, including pivots stored on one worksheet and sourced from another; appending rows to a source table expands matching PivotTable sources automatically; Refresh Summaries refreshes every saved report across the workbook and protects occupied output cells, and charts created from a summary track its refreshed range. Native XLSX PivotTable metadata, including supported percentage modes, including native x14 parent-total and percent-running-total calculations plus standard % Of, running totals, Difference From, rank, and Index calculations, per-field value-sort scope, caption filters, and Top/Bottom AutoShow filters, is written for this bounded layout.
- Refresh saved summaries manually or after ordinary edits, paste, undo and redo. Reports are generated as standard worksheet cells and can use existing chart features.
- GridForge provides persistent on-sheet slicer cards for ungrouped PivotTable row and column fields; slicer selections refresh the summary and export as PivotField item filters. Native Excel slicer objects and timelines, multi-level date hierarchies, calculated fields, Power Query M, Data Model, Power Pivot, DAX, OLAP and full Excel PivotTable editing remain unsupported. Summary-linked standard DrawingML charts export with native PivotChart links, although GridForge edits them as regular charts; specialized ChartEx charts do not carry native PivotChart links. A selected date field can be grouped by one interval (year, quarter, month, or day), and a numeric field can be grouped into custom intervals; both round-trip through supported native XLSX PivotTable cache metadata. GridForge query refresh is limited to worksheet transformations and public HTTPS CSV/JSON sources.

See the [v11.12 feature matrix](docs/FEATURE_VERIFICATION_V11_12.md), [release notes](docs/RELEASE_NOTES_V11_12.md) and [continuation guide](docs/CONTINUATION_GUIDE_V11_12.md). The 507 local or workbook-context function names do not establish exact Excel behavior, and 9 service-dependent names are not fully available offline.

## v11.11 highlights

- Move worksheet tabs left or right and set a tab color; colors persist in GridForge JSON and native XLSX.
- Hide or unhide selected rows and columns; visibility persists in GridForge JSON and native XLSX.
- Sort a selected table range by up to three ordered keys, with independent ascending or descending order per key.
- Find text or formulas across all visible worksheets and jump from results to the matching sheet and cell.
- Restrict cell entries by date or text length with the supported comparison operators; validation rules export to native XLSX.

Find All spans visible worksheets; Replace All remains limited to the active worksheet. Date input accepts ISO dates such as `2025-12-31` or Excel date serials. Text-length validation counts the entered text.

See the [v11.11 feature matrix](docs/FEATURE_VERIFICATION_V11_11.md) and [release notes](docs/RELEASE_NOTES_V11_11.md).

## v11.10 highlights

- Apply dropdown lists from typed options or a live range on the current worksheet.
- Restrict whole numbers and decimals with between, not-between, equality and comparison operators.
- Configure input prompts and stop, warning or information behavior for invalid entries.
- Save supported validation rules in GridForge JSON and native XLSX Data Validation records.
- Split one-column text into adjacent fields using comma, semicolon, tab, space or a custom delimiter, with a preview and quoted-field support.
- Hide and unhide worksheets; keep at least one worksheet visible and persist visibility in GridForge JSON and XLSX.

Inline lists are limited to 50 values / 255 total characters. Source ranges must be on the same worksheet; imported list validations can resolve direct defined names on that sheet and retain multiple target areas; conditional-format rules also preserve multiple areas. Custom-formula validation is not supported yet.

## v11.6 highlights

- Define, edit, delete and reuse worksheet-scoped names for cell references and ranges.
- Create names from selected heading rows or columns, and use names directly in formulas such as `=SUM(NetSales)`.
- Named range changes recalculate dependent formulas and support undo/redo.
- Names persist in GridForge JSON and the supported worksheet-scoped Excel `definedNames` subset in XLSX/XLSM.
- Retains the v11.5 Excel-style Formula ribbon: searchable function insertion, categorized function menus, SUM/AVERAGE/COUNT/MAX/MIN shortcuts, trace precedents/dependents, Show Formulas and formula inspection.
- Retains the existing chart, table, formatting, import/export, macro preservation, and desktop startup improvements from earlier releases.

Named ranges are worksheet-scoped and refer to cells or rectangular ranges. Workbook-scoped names and names defined as constants or arbitrary formulas are not included. GridForge remains a standalone application and does not claim full Excel feature parity.

GridForge uses Excel 2019 as an interaction reference while keeping its own UI and calculation engine. It does not embed, automate, or require Microsoft Excel.

## Download and run the compiled application

The compiled browser application is checked into `dist/`. To run it without building the TypeScript source, install Python 3.10 or newer and from the project folder run:

```powershell
python run_gridforge.py
```

Then open the local address printed in the terminal. On Windows, `START_HERE_WINDOWS.bat` starts the same precompiled app. The built native Windows installer is `build/release/GridForge-Setup-11.26.0-win64.exe` after the Windows build finishes; it contains the compiled desktop app and optional VBA bridge. End users need Windows 10/11 x64 and Microsoft Edge WebView2 Runtime. They do not need Node.js, npm, or Python when using the installer.

### Install using the included Windows installer

In the GitHub repository ZIP, open the `release/` folder and run `GridForge-Setup-11.26.0-win64.exe`. Follow the setup prompts and launch GridForge from the Start menu or its desktop shortcut. Windows 10/11 x64 and Microsoft Edge WebView2 Runtime are required. The installer includes the compiled application and optional Excel VBA bridge; Node.js, npm, and Python are not needed to install or run it. The installer does not install Microsoft Excel.

## Software dependencies

| Task | Required software |
|---|---|
| Run the precompiled browser app | Python 3.10+ (standard library only) and a current browser; no packages are installed by `run_gridforge.py`. |
| Edit and compile the frontend | Node.js 20+ with npm; dependencies are pinned through `package-lock.json` (`typescript` and `@xmldom/xmldom`). |
| Build Windows desktop binaries | Windows 10/11 x64, Node.js 20+, Python 3.12 x64 recommended, pip/network access for the pinned requirements in `desktop/requirements-windows.txt`, and Inno Setup 6. |
| Run the installed desktop app | Windows 10/11 x64 and Microsoft Edge WebView2 Runtime. |
| Run optional VBA bridge | Desktop Microsoft Excel and the bridge executable. GridForge itself does not require Excel. |

## Build and develop from source

From the project root in PowerShell:

```powershell
npm ci
npm run check
npm run build
python run_gridforge.py
```

`npm run dev` performs a frontend build and serves it locally. `npm run start` serves the already compiled `dist/`. The default source tree includes `dist/`, so Python-only use does not need Node. Do not use `npm run dev` if you want to avoid rebuilding the compiled output.

### Build the Windows installer

Install the builder dependencies above, then run:

```powershell
./desktop/build_windows.ps1 -SkipTests
```

The `-SkipTests` option builds the frontend, portable desktop app, VBA bridge, and installer without running the Windows test phases. Outputs go to `build/portable/` and `build/release/`. For a portable build without Inno Setup, use `./desktop/build_windows.ps1 -SkipTests -PortableOnly`.

### GitHub source and releases

The GitHub repository ZIP is `build/release/GridForge-v11.26.0-GitHub-Repository.zip`. It contains the project source, compiled `dist/`, lockfile, license, build files, GitHub Actions workflows, this README, and the Windows installer under `release/`. Extract it to start from a complete repository snapshot, or copy its root contents into a GitHub repository. The `Windows desktop release` workflow can build a Windows installer and portable ZIP from a `desktop-v*` tag or manual Actions run. The workflow performs its configured Windows checks; the local build command below does not run tests. The installer and archive SHA-256 checksums are listed in `build/release/SHA256SUMS.txt` next to the download files. No GitHub repository URL is configured in this working copy.

After the Windows desktop builder has created the installer and portable app folder, regenerate the portable ZIP, clean source archive, GitHub repository ZIP (including the installer), and checksums with:

```powershell
python scripts/package_source_release.py
```

The script packages the current portable app and bridge, plus source and compiled `dist/`; it excludes local build outputs and Python caches, then updates `build/release/SHA256SUMS.txt` for the installer, portable ZIP, and source archive.

## Example workbook

- `examples/GridForge_v11_2_Excel_UI_Chart_Tools_Demo.json` — calculated table column, data bars, dimensions and a combo chart with per-series error bars.
- `examples/GridForge_v11_2_Excel_UI_Chart_Tools_Demo.xlsx` — the example using native OOXML table, chart, conditional-formatting and dimension records.

Rebuild the v11.2 examples with `npm run demo:v112`.

## Verification

```powershell
npm run regression
python -m unittest discover -s tests -p "test_desktop*.py"
python scripts/smoke_formula_ribbon_v115.py
python scripts/smoke_defined_names_v116.py
python scripts/smoke_migration_v110.py
python desktop/gridforge_desktop.py --check
```

Previous versions: [v11.10 feature matrix](docs/FEATURE_VERIFICATION_V11_10.md), [v11.10 release notes](docs/RELEASE_NOTES_V11_10.md), [v11.9 feature matrix](docs/FEATURE_VERIFICATION_V11_9.md), and [v11.9 release notes](docs/RELEASE_NOTES_V11_9.md).






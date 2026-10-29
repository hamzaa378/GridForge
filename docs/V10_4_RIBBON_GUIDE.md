# GridForge v10.4 — Excel-style ribbon guide

The main worksheet opens in a quiet light-mode Home view. The menu tabs control which tools are displayed; functions and advanced analysis have not been removed.

| Tab | Choose it to… |
| --- | --- |
| File | Open/import Excel files or export XLSX/CSV/JSON and supported XLSM edits |
| Home | Format text, number formats, alignment, colors, and undo/redo |
| Insert | Insert AutoSum or add a worksheet using the + button by sheet tabs |
| Formulas | Browse functions using the right side panel; includes a separate External & Connected category |
| Data | Sort, filter, apply conditional rules, open Data Studio |
| Review | Find/replace or restore earlier edits |
| View | Freeze headings, show optional Insights |
| Automate | Record/replay GridForge actions or use the opt-in Windows Excel VBA bridge |
| More | Show every advanced action without hunting through the tabs |

## Shortcuts

`Ctrl+F1`: collapse/expand the ribbon. Left/right arrow keys while focusing a tab: switch to its neighbor. `Ctrl+O`: open; `Ctrl+S`: XLSX export; `Ctrl+Z` and `Ctrl+Y`: undo/redo; `Ctrl+F`: find; `F2`: edit the active cell; `Alt+=`: AutoSum. The formula bar is always visible. Light/dark and Guided/Excel workspace preferences are retained.

The **Insights** panel is closed initially to maximize visible spreadsheet space. Click it for formula dependencies, inspector, impact analysis and advanced statistics. Closing it does not change the workbook.

## Migration example

1. Open GridForge and choose **File → Open file** to import an XLSX.
2. Click a cell and type a literal value or `=SUM(B2:B10)`.
3. Use **Home** for formatting, **Data** for sorting/filtering, and **View** to freeze headings.
4. Use **File → Save Excel** to export, or save JSON to preserve GridForge-specific settings.
5. If a formula needs external service support, check its **External & Connected** note and evaluate it in a connected copy of Microsoft Excel instead.

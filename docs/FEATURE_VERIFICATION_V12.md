# GridForge 12.0.0 Feature Matrix

| Area | Status | Current implementation and limitation |
|---|---|---|
| Excel 2019 visual preset | Implemented | Familiar green-accent ribbon presentation; supports either light or dark theme. |
| Modern visual preset | Implemented | Teal/blue accent palette with matching ribbon colors; presentation only. |
| Compact visual preset | Implemented | Reduces top bar, ribbon, formula bar, tabs, and cell padding. Very small displays still use the responsive layout rules. |
| High Contrast visual preset | Implemented | Offers high contrast in both light and dark modes, strong keyboard focus outlines, and visible selected ribbon/cell states. It is not a certified accessibility conformance claim. |
| Workspace choice | Implemented | Excel-style and Guided workspace selection remains independent from appearance and theme. Guided mode keeps beginner entry points available. |
| Preference persistence | Implemented | Visual preset, color theme, and workspace are stored separately in the local application profile and do not modify workbook files. |
| `TOCOL` / `TOROW` ignore options | Improved | Referenced-cell errors are ignored for options 2 and 3; blank values are removed for options 1 and 3. Invalid option values return `#VALUE!`. |
| Specialized ChartEx labels | Improved | Value/category/series visibility, supported label position, and custom separator now round-trip through XLSX. ChartEx percentage labels and per-point positioning remain unsupported. |
| Conditional-format data bars | Improved | XLSX export now preserves the modeled solid/gradient fill and optional negative-value fill color. Excel's axis, border, direction, and threshold variants remain unsupported. |
| Large workbook and CSV refresh operations | Improved | XLSX decompression/checksum verification, cell updates, and formula evaluation yield between batches; CSV imports and public CSV query refresh yield while parsing. Local and cloud workbook opening report progress. Worker-imported and restored workbooks defer duplicate formula recalculation on the UI thread and calculate requested cells lazily. Routine autosaves use IndexedDB instead of synchronous local-storage serialization/writes, with local-storage migration/fallback; workbook snapshot construction remains on the UI thread. No-Worker import fallbacks remain synchronous. XML DOM parsing remains synchronous; workbooks remain subject to GridForge's supported row, column, file-size, and cell-count limits. |
| Cloud workbook access | Carried forward | OneDrive, Google Drive, and Dropbox are available in the Windows desktop build. Provider registration and scope limits remain; see `CLOUD_STORAGE_SETUP.md`. |
| Excel feature parity | Partial | GridForge remains a standalone application with known gaps in formulas, charts, PivotTables, external data, and advanced XLSX objects. No full parity claim is made. |

## Code and package verification

- `npm run check` — passed.
- `python -m py_compile desktop/gridforge_desktop.py` — passed.
- `npm run build` — passed during Windows packaging.
- Windows installer build — passed with `-SkipTests`.
- Automated tests, Windows launch, live UI, and Excel GUI checks — not run.


| Chart legend placement | Improved | Adds left placement to the existing top, bottom, right, and hidden options, including supported XLSX import/export. Some specialized chart rendering rules remain chart-specific. |

| XLSX import responsiveness | Improved | Local, cloud, and returned XLSM workbook parsing runs in a dedicated module worker; progress is relayed to the UI. XML parsing no longer blocks the main thread, and restore avoids recalculating all worker-calculated formulas on the UI thread. No-Worker fallback is synchronous; import size and worksheet limits remain. |

| LAMBDA formulas | Improved | Direct calls and callbacks for MAP, BYROW, BYCOL, MAKEARRAY, REDUCE, and SCAN run in scoped formula contexts; ISOMITTED tracks omitted parameters. Worksheet-local named LAMBDAs can be called directly or passed to supported lambda-aware functions; recursive definitions and workbook-global lambda names remain unsupported. |

| Formula-based GROUPBY/PIVOTBY | Improved | Adds LAMBDA aggregators, filter arrays, sort order, grand totals, PERCENTOF total bases, and multiple row-field columns for GROUPBY. PIVOTBY supports multiple measure columns, up to eight row and column field levels, hierarchy headers and field/value sorting, plus column, row, grand, parent-column and parent-row aggregation bases. Nested subtotals remain unsupported. |

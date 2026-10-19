# GridForge v2 — Excel-style feature specifications

## Sheet tabs

`SpreadsheetDocument` maintains an ordered list of up to 12 named worksheets, each backed by its own `Workbook` instance. Add, rename, duplicate and delete operate on document metadata. Duplicate uses the workbook JSON serialization to copy both cell content and formatting. Switching tabs retains each sheet's independent dependency graph and undo history until reload. The workbook title is editable in the top-left header. Duplicate names (case-insensitive), unsupported characters (`\\ / ? * [ ] :`) and empty names are rejected. The last sheet cannot be deleted.

This is NOT the same as full Excel workbooks. Formulas cannot use `SheetName!A1` and sheets are independent calculation domains.

## Cell styling

`CellStyle` is a sparse record with optional bold/italic/underline, `#RRGGBB` text and background, left/center/right alignment, and numeric format. Only valid CSS-safe six-hex-digit colors may be loaded from a JSON file. Styles do not affect dependency calculations and can be undone/redone. `General`, `Number`, `Integer`, `USD`, and `Percent` formats change the rendered value, not the raw double-precision numeric value or formula result.

Example: a stored value of `0.125` displays as `12.5%` when Percent is selected and as `0.125` under General.

## Copy / paste

Native Ctrl+C copies displayed TSV text so other applications can paste values. When the browser preserves a custom clipboard MIME type, the same GridForge tab can paste the raw formulas and source styles; relative references are shifted to the destination using the original row/column offset. Clipboard handling depends on the browser and its permissions; fallback is plain TSV with pasted text treated as cell source. Paste values uses the most recent **internal** copy snapshot and inserts **calculated values only**, without formatting; the text may be reinterpreted as a formula if it begins with `=` (a limitation of source-based editing). The raw cell and formatting paste is one undoable transaction.

## Range sorting

Selection must cover at least two rows. Both sort directions use the first (leftmost) selected column as a key. Numeric data is ordered numerically, text uses locale-aware comparison, and blank/error values go last. Every value within the selected rectangle travels as part of its complete row; cells outside the selection do not move. Formulas in moved rows shift relative references by the row displacement (locked `$` references remain locked). This policy is deliberately simpler than the full Excel sort/rewrite behavior. Always review results when sorting formulas that point to other sorted rows.

## Find / replace

Ctrl+F opens the Find/Replace dialog for the active sheet. Search can be case-sensitive; results are clickable addresses. Replace All excludes formula-source text by default to reduce unintended reference corruption. Including formulas is an explicit checkbox. Replace All creates a single undoable group of cell edits and asks for confirmation. It uses literal substring matching rather than regular expressions.

## Data Quality Lens and quick chart

The Analyze sidebar computes statistics for the selected rectangle: cells, numeric values, sum, average, minimum, maximum, blanks, formula errors, and exact duplicate values. A small, inline bar chart visualizes up to 16 points, taking labels from the first selected column and numeric values from the second selected column (or the first, if the selection is one column wide). It is intended for lightweight interpretation without importing a charting framework, and is not stored as a separate workbook object. Selection analysis reads cached results where possible.

## File formats and migration

- v2 workbook JSON shape: `{ format:'gridforge', version:2, activeId:'s1', sheets:[{id:'s1',name:'Sheet 1',workbook:{version:2,rows,cols,cells,styles}}] }`.
- v1 single-sheet workbook JSON (`{version:1,rows,cols,cells}`) can still be imported and becomes a one-sheet v2 document.
- Browser migration checks the v2 localStorage key first, otherwise the former v1 key, then generates sample data.
- CSV import becomes a one-sheet document. CSV export writes **calculated values** from the active sheet only; it cannot preserve formulas or styles.
- `.xlsx`, `.xlsm`, and `.ods` are not supported.

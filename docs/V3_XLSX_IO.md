# v3 XLSX interoperability: format, scope and safety

## Export flow

Toolbar **XLSX** -> `exportXlsx(doc)` -> `zipWrite(parts)` -> downloadable `.xlsx` Blob. The ZIP is **uncompressed (method 0)**, which is valid in standard OOXML spreadsheet files. Writing uncompressed avoids adding a third-party compression dependency. Excel, LibreOffice and Google Sheets support standard OOXML archives. Recalculation in Excel is requested using `<calcPr fullCalcOnLoad="1"/>` because a custom GridForge formula may not have an exported cached value.

Output ZIP parts include the content-types manifest, root relationships, workbook, workbook relationships, styles and one sheet XML per worksheet. Cells serialize as Excel numeric, boolean, inline string, formula, or styled blank. Style index maps are generated from sparse GridForge formatting. Native supported conditional rules use `<conditionalFormatting>` plus `<dxfs>`. Filter ranges produce `<autoFilter>`; freeze splits use `<pane state="frozen">`.

**Data fidelity:** original raw formulas are exported (without the leading `=` in `<f>`); supported formatting is serialized; only calculated numeric formula caches are emitted (Excel recalculates remaining formulas). Strings beginning with `=` are explicitly serialized as inline text unless the cell is a formula. Values and formulas are kept in each worksheet's source coordinate.

## Import flow

Toolbar **Import** → `.xlsx` → `zipRead(bytes)` (validate central directory, CRC-32, size/path bounds; support ZIP stored and raw DEFLATE) → `DOMParser` for XML parts → `SpreadsheetDocument` and `Workbook` objects → formula evaluator from the existing custom engine. The importer supports standard shared strings, inline strings (including rich text concatenation), numeric and boolean cells, basic style mapping, shared-formula reference shifting, sheet titles, active tab, filter ranges, frozen panes and simple cell-is conditional rules. It does not execute embedded formula macros, links, VBA or scripts.

**Limits:** up to 25 MB input archive, a 25 MB per-part uncompressed bound, 80 MB aggregate declared uncompressed size, 3,000 ZIP entries, 12 worksheets, 1,000 rows × 104 columns per sheet, 50,000 represented cells per imported sheet. ZIP64 and encrypted workbooks are unsupported. Oversized/nonstandard XML is rejected or raises a report. Do not import files from untrusted origins when browser memory is low.

**Import report:** The UI displays warnings on potentially lossy features. Unsupported formula syntax remains in the raw cell formula and is shown as a diagnostic (`#PARSE!`, `#NAME?`, etc.). Sheets with out-of-range populated cells are rejected rather than truncated. Malformed inputs leave the current document intact. When exporting after an import, only supported content can be preserved. Keep a separate original Excel file.

## Deliberate compatibility limits

- `Sheet2!A1` and quoted cross-sheet formula references are not supported by GridForge's formula parser; the app does not claim to recompute them. Excel may evaluate them on opening, but GridForge will show a parser diagnostic.
- Excel Tables (structured references), pivot tables, saved charts, images, merged cells, hyperlinks, comments, rich formatting, named ranges, data validation lists, date/time number-format fidelity and formulas with semicolon separators are not round-tripped.
- Unsupported conditional formatting types and compound filters are simplified/ignored. v3 UI supports one substring filter for a selected table column, and numeric greater/less/equal highlight rules.
- Workbook JSON v3 retains GridForge-specific filter criteria and formatting. A reimported XLSX only retains the compatible represented subset.
- Browser-native `DecompressionStream('deflate-raw')` is required for compressed XLSX; tested in local Chromium. If absent, import shows an explicit message.

## Reproducible tests

```bash
npm test                # includes ZIP CRC, OOXML export structure, filter/rule/JSON tests
npm run benchmark
# optional if Chromium and Python Playwright are installed:
GRIDFORGE_CHROMIUM=/usr/bin/chromium python scripts/browser_v3_smoke_optional.py
```

The browser test exports a workbook, confirms standard ZIP/OOXML parts, recompresses the archive using Python `zipfile.ZIP_DEFLATED`, imports it via the application's file picker, and compares formulas, sheet count, styles, frozen panes and conditional formatting. This demonstrates DEFLATE read, not merely re-opening the program's own stored ZIP file. Also manually open `examples/sample_sales.xlsx` in Excel or LibreOffice to assess UI rendering on your installation; that external-app check is not reported as completed here.

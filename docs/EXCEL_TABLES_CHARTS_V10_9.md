# Excel-inspired Table + Chart system — v10.9

## Implementation and source files

| Module | Responsibility |
|---|---|
| `src/engine/tables.ts` | Bounded table type, strict range/name validation, 10 built-in style names, cell-level header/band/first/last visual overlays, unique header extraction |
| `src/engine/sheettools.ts` | Validates/stores table definitions alongside charts, conditional rules, freeze panes and filtering |
| `src/engine/charts.ts` | Thirteen chart kinds, source-linked numeric series, SVG drawing, six chart palettes, schema validation |
| `src/engine/ooxml_charts.ts` | Produces native OOXML DrawingML chart parts and cached range references, including radar, percent-stacked and combined charts |
| `src/engine/xlsx.ts` | Native table parts, worksheet/drawing relationships and content-types; imports known table styles and supported chart types |
| `src/main.ts` | Excel-like ribbon, Home style/color shortcuts, Format as Table gallery, contextual Table Design actions, Insert → Chart dialog |
| `public/styles.css` | Lightweight table and swatch UI, light/dark consistency |

## File format

GridForge JSON schema `version: 3` stays backward compatible. `sheets[].tools.tables` is an optional list of table definitions. Each entry contains `id`, `name`, `range`, `style`, `header`, `bandedRows`, `bandedColumns`, `firstColumn`, `lastColumn`.

The `.xlsx` package contains `/xl/tables/tableN.xml`, `/xl/worksheets/_rels/sheetN.xml.rels`, per-sheet `<tableParts>` and the appropriate `[Content_Types].xml` entries. Named tables keep their selected built-in Office style string, column names derived from worksheet header values, header row setting and banding flags. Drawing relationships are preserved when a sheet has both charts and tables. Chart data references point to worksheet cells and cached numeric values.

**Security:** Table names, references and style options are validated before storage. OOXML import rejects external relationships, remote asset loading, entity declarations and invalid paths; importing any workbook never executes macros. GridForge macros run through their existing bounded native executor.

## Compatibility matrix

| Feature | Implementation status |
|---|---|
| Create and edit named tables | GridForge UI + JSON, native XLSX export/import |
| 10 built-in table styles | UI palette + JSON and native tableStyleInfo |
| Header, banded rows/columns, first/last column emphasis | UI, GridForge rendering, native style flags |
| Chart from selected table | Yes, uses table's current rectangular range |
| Live refresh when existing chart source cells change | Yes |
| Auto extend charts when appending rows to a table | Not yet |
| 13 chart kinds and 6 palettes | SVG GridForge rendering; native chart parts on XLSX export |
| Round-trip (app → XLSX → app) | Chromium verified for a combo chart and formatted table; 13-chart/11-table gallery imported successfully |
| Real Excel chart rendering/editing | **Not verified in this environment**; Windows COM script provided |
| Rich text within a cell, advanced conditional styles, structured references | Not implemented |
| True Excel PivotTables/PivotCharts | Not implemented |

## Testing

`npm test`: 889/889 passing, including nine new focused v10.9 tests. `python scripts/smoke_tables_charts_v109.py` exercises the actual formatting gallery, Table Design chart creation, font controls, color swatches, saving `.xlsx` and re-importing supported objects; passes in in-memory Chromium. `python -m unittest discover -s tests -p 'test_desktop*.py'`: seven desktop-wrapper tests pass. The 10,000-formula dependency benchmark passes. Native XLSX demo ZIP/XML integrity was also verified.

**Windows follow-up:** After building the installer and with Excel + pywin32 installed, run `python scripts/verify_excel_migration_windows_v109.py`. Confirm 13 chart objects and 11 ListObjects across two sheets. Review a sample of combo, radar and percentStacked charts visually, plus the Table Style gallery, because counting objects cannot establish complete visual equivalence. Repeat export/import with a small customer dataset and verify formatting, column labels and changing values.

## References

Design behaviors were taken from Microsoft's public guidance on *Create and format tables*, *Format an Excel table*, *Using structured references with Excel tables* and *Update the data in an existing chart*. This implementation is independently designed and does not copy Microsoft's application assets or implementation.

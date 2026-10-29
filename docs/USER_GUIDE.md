# GridForge v3 user guide update

**For v3 XLSX import/export, table filters, freeze panes, conditional formatting and AutoSum, start with the README or `docs/V3_ACCEPTANCE_CHECKLIST.md`. The instructions below are retained from v2.1 for the unchanged features.**

# User Guide

## Opening the app

Use `npm install` and `npm run dev`; open `http://localhost:4174`. GridForge loads a sample sales workbook with example formulas, totals and a product lookup. Everything runs in your browser. The original 20,800-cell address space is much larger than the cells visible on screen.

## Editing a cell

Single-click a cell to select it; double-click to edit its raw content. You can also type immediately to replace the current cell, or press Enter/F2 to edit it. Press Enter to commit (then move down), Tab to commit (then move sideways), or Escape to cancel. The formula bar above the grid can also be used: select a cell, type something beginning with `=`, press Enter or click ✓.

A raw number like `250` becomes numeric. `TRUE`/`FALSE` becomes boolean. Everything else becomes text. Prefix a leading apostrophe to force literal text: typing `'=SUM(1,2)` displays `=SUM(1,2)` without evaluation.

## Selecting / filling / navigation

| Action | Shortcut / gesture |
| --- | --- |
| Move by one cell | Arrow keys |
| Extend a selection | Shift + arrow keys |
| Jump to a cell | Enter `AA120` in address box, press Enter |
| Edit cell | Double-click, F2, Enter, or type |
| Commit / cancel | Enter or Tab / Escape |
| Copy values externally, raw formulas/styles within GridForge | Ctrl+C (or Cmd+C) |
| Paste GridForge formulas/styles or external TSV | Ctrl+V (or Cmd+V) |
| Fill selected range down | Ctrl+D |
| Drag formula down or across | Drag purple handle at bottom-right of active cell |
| Clear range | Delete / Backspace |
| Undo / redo | Ctrl+Z / Ctrl+Y |
| Jump to address input | Ctrl+L (when grid focused) |
| Find / replace | Ctrl+F or toolbar Find |
| Bold / italic / underline | Ctrl+B / Ctrl+I / Ctrl+U (grid focused) |
| Paste calculated values | Copy selection, choose target, click Paste values |
| View all shortcuts | Shortcuts in top header |

Drag-fill shifts **relative** references by destination offset. `=$A$1+B2` dragged one row down becomes `=$A$1+B3`.

## Multiple worksheets and formatting (v2)

Use the `+` button next to the sheet tabs to add a worksheet. Click a tab to switch sheets. Rename, Duplicate and Delete affect the active sheet. A duplicated worksheet copies all raw formulas and cell formatting into an independent sheet. The last remaining sheet cannot be deleted; cross-sheet formulas are not supported. Click the workbook title at the top to rename the document.

To format a range, select one or more cells, then use the second toolbar: Bold, Italic, Underline, number format (General, Number, Integer, USD, Percent), alignment, text/fill colors, borders, or Clear style. Border presets include all four edges, or one individual edge. Enter an Excel custom number-format code in **Format code**, for example `0.00;[Red](0.00)` or `$#,##0.00`. GridForge keeps the code in XLSX; the on-screen renderer handles a common subset of numeric placeholders, grouping, decimals, percentages, literals, colors, negative/zero sections, and scientific notation. A fourth text section can use `@` to insert the cell text. Unsupported conditions and tokens fall back to the standard display. Number formatting is non-destructive: `0.125` displays as `12.5%` but still evaluates as numeric `0.125`. Undo and redo also work for style changes.

For conditional formatting, choose a two-color or three-color scale; the optional midpoint uses the 50th percentile. Numeric rules include greater/less than, inclusive bounds, equality/inequality, and between/not-between. GridForge also has data bars, but Excel icon sets, formula-driven rules, and all data-bar extension settings are not yet fully represented in XLSX.

To sort, highlight a rectangular range covering two or more rows; use Sort A–Z or Sort Z–A. The leftmost selected column is the sort key. Only the selected rectangle moves; include every related column you want to keep paired with its row. You can undo the sort. For Find/Replace, open Ctrl+F or the toolbar button; replacements exclude formula-source text by default.

To inspect data, select e.g. A4:B11, then open the **Analyze** sidebar. See numeric counts, sums, averages, blanks, formula errors, duplicate values and an inline quick chart. The chart is a view, not a saved graph in the JSON workbook. Formula Inspector, Impact and What-if continue to work as before.

## Inspector (default sidebar)

Select cell **D4**. Its result is `4000`, and the inspector shows the expression tree for `=B4*C4`, the input addresses, and current values. Click an input chip to navigate to that cell.

If a cell has an error, the panel provides the code, explanatory message, formula with a caret, and a suggestion. Try `=SUM(A1, B2` in I25 to see the missing-closing-parenthesis diagnostic.

## Impact panel

Select B4 and click **Impact**. The panel lists direct upstream inputs and reachable downstream dependents. These are based on the real dependency graph used by the evaluator. The links jump to cells and let you inspect a calculation path.

## What-if panel

Select B4 and click **What-if**. Enter a hypothetical value of `300`, then press **Preview impact**. The panel lists which cells would change (such as D4, F4, G4, and totals) but **does not modify B4**. If the proposed change is desirable, click **Apply this change**. The application commits it as one undoable edit.

## Functions panel

Select **Functions**, search `VLOOKUP`, or scroll the 56 function cards. Each card contains a description and example. Click one to insert its function template into the formula bar. Enter the arguments and press Enter to commit. The complete reference is in [FUNCTION_REFERENCE.md](FUNCTION_REFERENCE.md).

## Files and persistence

- **Demo** reloads the original sample. It requests confirmation because it overwrites the current workbook.
- **New** starts a blank v2 document with one 400×52 worksheet after confirmation.
- **Import** accepts `.csv` and `.json` files from your local machine.
- **CSV** exports calculated cell values into a spreadsheet-friendly CSV. Formulas become their output values. Formula-like text is escaped with a leading apostrophe for safer external opening.
- **JSON** exports all worksheet names, formulas, dimensions, formatting, and selected worksheet. Use this as a reliable backup, and re-import it to continue working.
- **Autosave** writes to browser localStorage after each edit, grouped into a small debounce. Reload to see the previous workbook again. Storage is local to this browser and site origin.

## Page setup and printing

Choose **Page Layout → Page Setup** to set portrait or landscape orientation, Letter/Legal/A3/A4/Tabloid paper, fit-to-page width and height, margins, gridlines, headings, and an optional print area. **Print Sheet** opens a print-ready browser view. The same basic settings are saved to XLSX. Excel-only options such as repeating title rows, manual page breaks, headers/footers, and print order are not included yet.

## Refreshable offline queries

On the **Data** tab, choose **New Query** to create a worksheet cleanup pipeline or an external CSV/JSON query. Worksheet queries use a source range; external queries use a public HTTPS URL and require the endpoint to allow browser CORS requests. Requests omit cookies and credentials, reject redirects, and have a 3 MB response limit. Select up to 16 cleanup/conversion steps and a separate load destination. **Refresh Queries** replays worksheet pipelines or fetches external sources again. GridForge documents keep the query definition and cached result. XLSX export includes the last fetched external result as ordinary cell values; it does not create an Excel data connection or Power Query M query. Authentication, database connectors, query folding and scheduled refresh are not supported.

## Sample data, examples and expected results

- D4: `=B4*C4` → 4000 at startup.
- F4: `=ROUND(D4*E4,2)` → 960 at startup.
- G4: `=IF(F4>1500,"On track","Review")` → Review at startup.
- D14: `=SUM(D4:D11)` → 34739 at startup.
- J5: product code `P-103`.
- K5: `=VLOOKUP(J5,A4:C11,3,FALSE)` → 120 at startup.
- K11: lookup of the highest-selling product using `INDEX`, `MATCH`, and `MAX`.

More formulas to try:

```text
=SUM(B4:B11)
=AVERAGE(C4:C11)
=IF(B4>100,"High","Low")
=COUNTIF(B4:B11,">100")
=VLOOKUP("P-103",A4:D11,4,FALSE)
=ROUND(PI(),3)
=IFERROR(1/0,"handled")
```

## Problems / troubleshooting

**The app page is blank:** Confirm Node.js 20+ is installed, run `npm install` and `npm run build` in the project folder, then `npm run start`. Inspect the browser console for errors. Do not double-click `index.html` from the file manager because ES modules use HTTP-origin loading.

**Default port already in use:** Stop the old server, or choose a different port. On Windows PowerShell: `$env:PORT=5050; npm run start`. On macOS/Linux: `PORT=5050 npm run start`.

**A formula shows #PARSE!:** Verify every `(` is paired with `)` and function arguments are separated by commas. Open Inspector for the caret and suggestion.

**A formula shows #CYCLE!:** The cell reads itself directly or through other formulas. Use Inspector and Impact to find and break the loop.

**Missing saved edits:** localStorage may be disabled, cleared, full, or isolated to another browser or port. Export to JSON regularly. A file that is imported does not sync across tabs or machines.

**CSV did not preserve formulas/sheets/styles:** This is intentional. Export JSON to preserve those. CSV import/export operates on one active sheet and values only.

**Some Excel formulas do not work:** Only the functions listed in [FUNCTION_REFERENCE.md](FUNCTION_REFERENCE.md) are implemented; support is intentionally narrower than Microsoft Excel.

## Pivot Summary

Choose **Insert → Pivot Summary** to summarize a headered worksheet range. Rows are optional, so you can create a Values-only report. Arrange up to eight nested fields in each of Rows and Columns, reorder Rows, Columns, and primary Values with the up/down controls, and choose up to eight value calculations (a source field may also appear in Rows or Columns, and you can add it more than once) with independent custom names and summaries (Sum, Count, Count Numbers, Average, Minimum, Maximum, Product, standard deviation or variance); new calculations default to Sum for numeric source fields and Count for text or error fields. Count includes formula cells even when their displayed result is an empty string, while Count Numbers includes only numeric results. Set each measure to show its value normally; as a percentage of the grand, row, column, parent-row, or parent-column total; as % Of a Previous, Next, or specific item in a selected Base field; as a running total in a Base field; as Difference From/% Difference From a Previous, Next, or specific item; or as rank smallest-to-largest/largest-to-smallest within a selected row/column Base field, or calculate Index from row, column, and grand totals, apply searchable multi-item report filters or hide selected row and column items, and filter row/column labels by equality, prefix/suffix, containment, caption ordering, or inclusive Between/Not between bounds. Filter lists are capped onscreen and can be narrowed by text search, so larger sources do not create thousands of dialog options at once. Blank categories are included. Blank source headers are named Column 1, Column 2, and so on; duplicate headers receive unique suffixes such as (2), keeping field choices and formula criteria unambiguous. A date range can filter an ungrouped row or column field with inclusive start and end dates; the range is saved with the Pivot Summary, and XLSX export represents it as hidden PivotField items rather than a native Excel Timeline. If a filter excludes every record, the summary keeps its field headers and shows requested grand totals as zero. Row/column intersections with no source records are blank by default; set **For empty data cells, show** to display a replacement such as `0` or `N/A`. The setting round-trips through supported native PivotTable XLSX metadata. Pivot output is capped at 250,000 cells to keep refreshes responsive. You can choose tabular (separate row-field columns), outline (grouped fields), or compact (indented labels) report layout, set a global sort default, then choose ascending, descending, or source order separately for each row and column field. Each selected row or column field can also be sorted by a selected Values calculation, ascending or descending; ties use ascending label order. Ordinary field sorting compares underlying numeric and date values, with text labels sorted using a locale-aware natural order. The Top/Bottom value filter limits each selected field to up to 255 items ranked by any selected Values calculation within its parent group. XLSX import and export preserve all three layout selections in the supported PivotTable subset. For a date field on Rows or Columns, choose one grouping interval (Years, Quarters, Months, or Days) in the field controls; non-date values prevent grouping. For a numeric field, enter a positive **Bin size** to group values into intervals; bounds default to source values, with optional Start and End overrides; intervals that would create more than 4,096 groups are rejected. Place row subtotals above or below each group, add optional blank rows after parent groups, set subtotals independently for selected row or column fields, and turn row or column grand totals on or off. Double-click a value, subtotal, or grand-total cell, or choose **Show PivotTable details** from its context menu, to create a **Details** worksheet containing matching source records after the report's active filters, with source cell formatting and custom row heights and column widths copied where available; the workbook's worksheet row limit also limits the number of detail records. Nested row and column groups show **+ / −** controls beside their labels; use them to collapse a parent to its subtotal or expand it to show child groups. When editing a saved Pivot Summary, use **Collapse all** or **Expand all** to change both axes together. The expansion state is saved with the workbook and retained during refresh. XLSX export writes the native collapsed-item flag and GridForge restores it on import where field-item labels can be resolved. Excel stores that flag on a field item, so repeated labels under different parents may share the collapsed state. This OOXML behavior is code-checked but has not been verified in a live Excel session. Select a cell in a saved summary and open Pivot Summary to create a linked chart, refresh, update or delete it; linked charts follow Pivot Summary range changes. Saved reports also refresh when their source is edited. You can select a formatted table directly as the source, or enter its exact range. GridForge stores the table name, follows its current range during refresh, and exports an XLSX PivotCache that points to the table by name; imported table-backed pivots keep that link. Renaming or deleting that source table leaves the pivot source unresolved. See [table-backed PivotTable sources](PIVOTTABLE_TABLE_SOURCES.md). Standard DrawingML summary-linked charts export with native PivotChart links, but remain regular editable GridForge charts. Specialized ChartEx chart types do not have native PivotChart links. Export writes a bounded native XLSX PivotTable definition and cache, including supported percentage, % Of, running-total, Difference From, and rank modes; parent-row, parent-column, percent-running, Difference From, rank, and Index calculations use native Excel PivotTable metadata. For a configured value-sorted field, XLSX export writes the native auto-sort scope and selected Values measure; import reads that scope back into GridForge. Top/Bottom filters use native PivotField AutoShow metadata and round-trip through supported XLSX files. Label filters round-trip as native PivotTable caption filters. GridForge slicer cards can filter an ungrouped row or column field, cache their item list until source values change, and save selected items as PivotTable field filters. One slicer can connect compatible summaries on the same worksheet when they use the same source range and expose the same ungrouped field; use its chain button to edit connections. Changing selection updates all connected summaries, newly connected summaries inherit the selection, disconnected summaries recover their prior field filters, and deleting the slicer restores prior filters for all connected summaries. Native Excel slicer cache/view parts now import and export with selected items and drawing placement. Pivots with the same source and grouping share an exported native cache, allowing a slicer to connect to those compatible pivots; if connected summaries require different native caches, the Excel slicer is limited to its primary cache. This XLSX behavior is code-checked but has not been verified in a live Excel session. Timelines, multi-level date hierarchies, calculated fields, and Data Model connections are not implemented.

`GETPIVOTDATA` reads a saved generated Pivot Summary by its value caption and optional field/item criteria, for example `=GETPIVOTDATA("Sum of Sales",A1,"Region","North")`. It applies the report's active report, item, caption, date, and Top/Bottom value filters and calculates the selected base aggregation and Show Values As mode. The pivot reference must point inside the generated report. Formulas recalculate when the report region changes. Results are based on GridForge's generated Pivot Summary settings and are not certified for every Excel PivotTable cache or subtotal combination.

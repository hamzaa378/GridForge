# User Guide

## Opening the app

Use `npm install` and `npm run dev`; open `http://localhost:4173`. GridForge loads a sample sales workbook with example formulas, totals and a product lookup. Everything runs in your browser. The original 20,800-cell address space is much larger than the cells visible on screen.

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

To format a range, select one or more cells, then use the second toolbar: Bold, Italic, Underline, number format (General, Number, Integer, USD, Percent), alignment, text/fill colors, or Clear style. Number formatting is non-destructive: `0.125` displays as `12.5%` but still evaluates as numeric `0.125`. Undo and redo also work for style changes.

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

**Port 4173 already in use:** Stop the old server, or choose a different port. On Windows PowerShell: `$env:PORT=5050; npm run start`. On macOS/Linux: `PORT=5050 npm run start`.

**A formula shows #PARSE!:** Verify every `(` is paired with `)` and function arguments are separated by commas. Open Inspector for the caret and suggestion.

**A formula shows #CYCLE!:** The cell reads itself directly or through other formulas. Use Inspector and Impact to find and break the loop.

**Missing saved edits:** localStorage may be disabled, cleared, full, or isolated to another browser or port. Export to JSON regularly. A file that is imported does not sync across tabs or machines.

**CSV did not preserve formulas/sheets/styles:** This is intentional. Export JSON to preserve those. CSV import/export operates on one active sheet and values only.

**Some Excel formulas do not work:** Only the functions listed in [FUNCTION_REFERENCE.md](FUNCTION_REFERENCE.md) are implemented; support is intentionally narrower than Microsoft Excel.

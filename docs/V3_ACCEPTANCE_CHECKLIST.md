# GridForge v3 — browser acceptance checklist

Run `npm install && npm run dev`, then open the URL printed by the server (normally `http://localhost:4174`). Use **Demo** or `examples/sample_sales.xlsx` for a controlled test. Work through this list on your own machine before submitting the GitHub URL.

| ID | Steps | Expected result |
| --- | --- | --- |
| A1 | Edit `B4` in Demo, inspect `D4` | Revenue recalculates immediately; Inspector shows the formula |
| A2 | Right-click a cell | Native-looking GridForge menu; Copy/Paste/Sort/AutoSum/Filter/Conditional actions appear |
| A3 | Enter `=SUM(B4:B11)` into empty cell | Numerical sum appears; an intentionally malformed formula gives a useful diagnostic |
| A4 | Copy formula from one cell to another | Relative references adjust, absolute references remain pinned |
| A5 | Select `A3:G11` → Custom Sort with header | Rows and styling move together, Undo restores state |
| B1 | Import `examples/sample_sales.xlsx` | Sales and Summary sheets appear; dismiss import notes |
| B2 | Sales → `D2` | Formula bar shows `=B2*C2`; result is 75000 |
| B3 | Sales → `D7` | Formula bar shows `=SUM(D2:D5)`; total equals 1,102,500 |
| B4 | Select `A1:D5` → Filter, column A, contains `Mouse` | Only header and matching row appear within table; Show all restores other rows |
| B5 | Select `D2:D5` → Conditional → greater than 100000 | D4 (960000) is highlighted and D3 (5000) is not |
| B6 | Click Freeze row and Freeze col, then scroll | First header row and A column stay pinned to viewport |
| B7 | Click blank `B6`, choose AutoSum | Formula sums contiguous numeric values in column B above; verify formula bar |
| C1 | Add/rename/duplicate a worksheet | Independent content and tab names persist |
| C2 | Change a cell's bold/fill/currency style, undo/redo | The exact style comes back with Redo |
| C3 | Click XLSX | A file with `.xlsx` extension downloads successfully |
| C4 | Open the exported XLSX in Excel/LibreOffice | Values, sheets, formulas and supported styles render; check calculated values; report any warnings |
| C5 | Import exported `.xlsx` into GridForge | Supported data, formula strings, styles, filter range and freeze state survive |
| C6 | Export JSON, reopen in a new origin, import JSON | Custom formula data plus all native GridForge v3 view/rule metadata survive |
| D1 | `npm run check` | Exit code 0 |
| D2 | `npm test` | 57/57 tests pass |
| D3 | `npm run benchmark` | A source edit invalidates only its affected dependents (26 in provided benchmark) |
| D4 | GitHub Action CI / Pages workflow | Build and deployment both show green only after actual GitHub run |

**Known limits for reviewers:** 12 sheets maximum; 400×52 default address space, 1000×104 max imported; up to 30 simple conditional rules; only one text-contains filter at a time per sheet; no cross-sheet formula evaluation; XLSX support is subset only. Do not describe unsupported Excel operations as implemented.

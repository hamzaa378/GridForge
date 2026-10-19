# GridForge v2.1 — Excel-inspired cell context menu and custom sorting

## Changes implemented

The HTML grid supports native browser right-click interception **only inside spreadsheet cells**, without disabling the ordinary context menu elsewhere in the application. Right-click opens a custom keyboard-accessible floating menu with these commands:

| Menu group | Commands | Notes |
|---|---|---|
| Editing | Edit cell, Cut, Copy, Paste, Paste values | Cut is staged and only removes the original when a same-sheet paste succeeds. Copy/Paste uses an internal typed snapshot; the browser's clipboard API is best-effort for Copy. |
| Cleanup | Clear contents, Clear formatting | Two separate undoable operations; deleting contents preserves formatting. |
| Format | Toggle bold | Uses the same style engine and undo history as the ribbon. |
| Sort | Sort A–Z, Sort Z–A, Custom sort | All selected columns are moved together. The custom dialog chooses any column **within** the selected rectangle and can exclude a header row. |
| Inspect | Find and replace, Formula Inspector, Analyze selection | Reuses existing functions and sidebar, not duplicate logic. |

If a user right-clicks **inside** an existing multi-cell selection, it remains selected. Right-clicking **outside** the selection selects the clicked cell. The menu clamps to the visible viewport, is dismissible by Escape and click-outside, and can also be opened with **Shift+F10** (when the grid is focused). Up/Down arrow keys move between menu commands.

## Sorting semantics and safety

Sorting is **stable**: if two keys compare equal, their prior order is retained. Numbers are compared numerically; text uses locale-aware comparison with numeric chunks; empty/error values come last. Moving rows also moves sparse `CellStyle` snapshots. Values and styles are written using a single `Workbook.applyTransaction`, so Ctrl+Z/Ctrl+Y restores both together.

**IMPORTANT:** a sort never expands selection to the rest of the table automatically. To preserve row integrity, select **all columns of the table** first, e.g. `Q1:S5` to include a header. Choose Custom sort and enable **My selection has a header row**. A sort of `R2:S5` intentionally leaves Q unchanged, just as the warning in the dialog explains.

Moved formulas have relative references translated according to the row displacement, with absolute `$` references preserved. This is the documented GridForge simplified formula-moving rule; complex Excel reference rewriting, structured tables, merged cells, and hidden/filter rows are not implemented. Always verify complex formula relationships after a sort.

## Copy / cut / paste rules

- **Copy** snapshots raw formulas, displayed values and styles. Same-app paste adjusts relative references from copied coordinates to destination coordinates and preserves styling. One operation = one undo step.
- **Cut** snapshots a range and defers clearing its original. Same-sheet paste moves it atomically, including its styles, and supports overlapping source/destination ranges. This is more reliable than deleting on Cut. Clipboard cut state clears after a successful move.
- If a Cut snapshot is pasted into **another** worksheet, it is treated as a copy; the original worksheet is never silently cleared. This avoids accidental cross-sheet data loss.
- **Paste values** inserts evaluated text without style; text beginning with `=` is escaped to remain literal rather than execute as a formula. It does not consume a pending Cut.
- A paste completely outside grid bounds is rejected without modifying cells.
- Context Paste currently prefers the most recent **internal** copied range. To paste externally copied text from Excel or another application, click a cell and use **Ctrl+V**. Pasting external TSV uses basic text parsing and does not import Excel formatting.
- Clipboard content lives only in the current application tab's memory; it is not shared across tabs or persisted after reload. Browser clipboard permissions may prevent copying to other apps via the context menu; use Ctrl+C for normal OS-level clipboard interaction.
- Cut/move does **not** automatically rewrite formulas in third-party cells pointing to the moved range; formula moves are a deliberately limited workflow. Check dependent formulas after moves.

## Reproducible sorting acceptance test

1. Fill Q1:S5 with the following data:

   | Product | Quantity | Price |
   |---|---:|---:|
   | Keyboard | 50 | 1500 |
   | Mouse | 10 | 500 |
   | Monitor | 80 | 12000 |
   | Webcam | 25 | 2500 |

2. Highlight Q1:S5 (include header), right-click **inside** the selection, choose **Custom sort**.
3. Select column R (Quantity), choose ascending order, check **My selection has a header row**, then sort.
4. Expect the product names in Q2:Q5 to be **Mouse, Webcam, Keyboard, Monitor** and the quantities R2:R5 to be **10, 25, 50, 80**; the Price must remain with its original product.
5. Check Q1:S1 is untouched and Ctrl+Z restores all original rows; Ctrl+Y reapplies the sort.
6. Apply Bold to the original Q3 (Mouse) before sorting. Confirm Bold moves to Q2 with Mouse, and Undo restores both formatting and rows.

## Reproducible Find/Replace acceptance test

1. In the unsorted dataset, press Ctrl+F; search `Mouse`, and click **Find all**. Expect Q3 (the address changes after sorting).
2. Replace with `Wireless Mouse` using Replace all and confirmation. Verify the displayed text.
3. Ctrl+Z restores the original. Search `mouse` with Match case unchecked (finds Mouse), then checked (no exact-case match).
4. Formula-protection check: `=UPPER(Q3)` stays a formula when Replace All is used with **Include formulas** unchecked.

## Regression and demonstration

- `npm run check` — TypeScript type checking.
- `npm test` — all engine tests, including sorting styles, overlapping cut/paste, header-aware sorting, cross-sheet safety, and source text escaping.
- `npm run benchmark` — 10,000-formula workload.
- `GRIDFORGE_CHROMIUM=/usr/bin/chromium python scripts/browser_smoke_optional.py` — optional Chromium smoke test (requires Playwright, installed separately). Exercises right-click, menu copy/paste, Shift+F10, Custom sort, undo and cut.
- `npm run dev` and then open the URL printed in the terminal. The server tries available ports starting at 4174 (through 4184) when `PORT` is unset; explicit `PORT` is honored.

The repository is ready to push, but GitHub Pages is **not** deployed or tested until you push and enable it. This is not a claim of complete Excel compatibility or XLSX support.

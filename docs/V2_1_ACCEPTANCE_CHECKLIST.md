# GridForge v2.1 — Manual browser acceptance checklist

Perform these checks on the new version, ideally in a new browser profile or after **exporting your saved JSON workbook**, so you do not overwrite existing data.

- [ ] Run `npm install`, `npm run check`, `npm test`, and `npm run dev`. Open the URL printed, which may be 4174 or 4175.
- [ ] Right-click a single populated cell; custom menu appears and stays fully within the viewport.
- [ ] Right-click the middle of a multi-cell selection; selection remains intact.
- [ ] Right-click outside that selection; the clicked cell becomes selected.
- [ ] Press Shift+F10 when grid focused; menu opens, arrows navigate, Escape closes.
- [ ] Copy D4 and paste to H4; raw formula changes from `=B4*C4` to `=F4*G4`.
- [ ] Undo that paste restores H4; Redo reapplies it.
- [ ] Cut a cell and paste elsewhere on the same sheet; source clears **only after** successful paste. Undo restores both source and destination.
- [ ] Copy text from outside the app and paste with Ctrl+V (the context Paste uses an internal snapshot).
- [ ] Copy a formula cell and use Paste values; destination contains evaluated result (not a formula).
- [ ] Use context Clear contents, undo, Clear formatting, undo.
- [ ] Sort a complete table by a non-first column using Custom sort, retain the header and each record's style. Undo/Redo is exact.
- [ ] Find and Replace matches the expected source string and does not modify formulas unless explicitly enabled.
- [ ] Data Quality and Formula Inspector open from the context menu.
- [ ] Export JSON, import JSON again, and confirm formulas and styles remain intact.
- [ ] Refresh the page after the save indicator changes to **Saved locally**, verify persistence in the current browser.
- [ ] No errors in browser DevTools Console during these actions.

Report PASS/FAIL for each item. Do not claim performance numbers, browser/storage compatibility or deployment without reproducing the tests in your target environment.

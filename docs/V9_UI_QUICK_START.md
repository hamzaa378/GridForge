# GridForge v9 — Excel user quick start

The top command tabs are designed around common Excel tasks. **All tools** retains the full legacy toolbar and is the initial tab, so nobody needs to learn the new interface to use the application.

## Three-minute tour

1. **Home:** Begin a new workbook, load the demo, and undo/redo an edit. Click `B4` and enter `200`; the dependent revenue cell updates through the custom formula graph.
2. **Formulas:** Select `Z10`, type `=SUM(B4:B11)` and press Enter. Or use Formulas → sidebar Functions, choose a category, and click a function card to insert its example. Search finds both names and descriptions.
3. **File:** Import an `.xlsx` or `.xlsm`, export as XLSX/CSV/JSON, or save supported edits back into a macro-enabled XLSM.
4. **Data:** Select a rectangle, sort/filter, or use Data Studio to preview and apply a cleanup step with Undo support.
5. **Automate:** Record a safe GridForge action macro or, on Windows with installed desktop Excel, run a **trusted** existing VBA procedure through the opt-in bridge.

Use **command search** at the top: type a task such as `sort`, `formula`, `import`, `filter`, `macro`, or `undo`, then press Enter. The relevant command group becomes active.

The **Start here** button opens an optional onboarding dialog. It does not appear automatically and never blocks a reviewer’s test run.

## Excel keyboard conventions

- Arrows: navigate; Enter/F2: edit a cell; Shift+Arrow: extend selection.
- Ctrl+C/V: copy/paste; Ctrl+Z/Y: undo/redo; Alt+=: AutoSum; Shift+F10: context menu.
- Right-click the grid for an Excel-inspired cell menu.
- Formula bar displays the raw source of the active cell, while its grid result displays calculated output.

## Known differences

- The default sheet is 400 × 52 (20,800 addressable cells) with virtualized rendering, rather than Excel’s full worksheet size.
- Advanced Excel formulas that spill arrays do not spill in GridForge. Use `INDEX` to select an element of an array-valued function.
- Other familiar Excel features such as filters or conditional formatting cover a documented subset.
- Workbooks with VBA need trusted desktop Microsoft Excel to **run VBA**, even though GridForge can preserve macros in supported `.xlsm` edits.

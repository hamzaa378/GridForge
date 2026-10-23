# v10 — light/dark themes and Excel-to-GridForge onboarding

## What changed

The top bar has two independent settings:

1. **Workspace** → **Excel familiar** (ribbon and full tools) or **Guided & simple** (beginner quick actions and Home toolbar). The two modes share the **same in-memory workbook and undo history**.
2. **☀ Light / ☾ Dark** switches the complete interface palette without altering data, formatting stored in the workbook, formula results or preferences other than the theme itself.

The preferences are saved at `gridforge:theme:v10` and `gridforge:workspace:v10` in `localStorage`. A small script in `public/index.html` reads them before the stylesheet loads so the initial paint matches the saved theme. Invalid or inaccessible storage defaults to dark + Excel familiar.

## Five-minute walkthrough for an Excel user

1. Start with **Workspace → Excel familiar** and whichever color theme is comfortable.
2. Click any cell or use the name box. Type a number or `=SUM(A1:A5)` in the formula bar, then Enter.
3. Use **Home** for Undo/Redo, **File** for import and export, **Data** for sorting and filtering, **Formulas** to search 506 registered functions.
4. Right-click a selection to copy/paste, clear, format, sort, run AutoSum, or inspect a formula.
5. Choose **Guided & simple** to see the starter buttons and short guide. Switch back at any time; you do not need to reload the page.

## Beginner workflow

1. Set **Guided & simple**.
2. Choose **Load example** to see formulas and totals. The app asks before replacing your active workbook.
3. Click **Browse formulas** or **Show tutorial**.
4. Use **Open Excel file** to import an existing XLSX, XLSM, CSV or JSON workbook.
5. Export JSON when finished to preserve as much GridForge-specific information as possible.

## Keyboard behavior

Arrows, F2, Enter, Ctrl+C, Ctrl+V, Ctrl+Z, Ctrl+Y, Ctrl+F, Shift+selection, and Alt+= follow the existing GridForge Excel-inspired behavior. Switching workspace or theme does not remap shortcuts.

## Manual acceptance tests on Windows

- Choose Light, reload the same origin: Light should remain selected and the first paint should be light (no dark flash).
- Choose Guided, reload: Guided should remain selected.
- Edit a cell in Guided, then switch to Excel familiar: its value and dependent formulas must remain unchanged.
- Undo after switching either setting; the last workbook edit should be undone normally.
- Select an existing cell with fill/font formatting; changing the app theme should not erase its stored style.
- Switch light/dark while an XLSX workbook is open, export, re-import and compare formulas.
- Test with a narrow window: ribbon and setting controls must remain usable via horizontal scrolling.

**Origin note:** Chrome/Edge storage is partitioned by hostname and port. `localhost:4174` and `localhost:4175` are different storage origins. Export JSON before moving workbooks between ports.

## Known limitations

A few rarely used legacy dialog details may retain darker accents in light mode. Custom cell colors are authored as explicit spreadsheet styles; their contrast is not automatically altered to match app themes. Browser privacy settings may prevent preference persistence.

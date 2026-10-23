# GridForge v10.3 — Excel / Google Sheets-inspired workspace

GridForge is not a visual clone of Microsoft Excel or Google Sheets. It follows familiar spreadsheet navigation while retaining its distinct identity.

## Clean starting view

- **Light appearance by default**, with a saved light/dark toggle. Existing users' stored theme settings are respected.
- **Home ribbon by default** rather than showing all actions at once. `File`, `Formulas`, `Data`, `Tools`, `View`, and `More` reveal appropriate controls.
- **Open file** and **Save Excel** are available on Home, with standard shortcuts **Ctrl+O**, **Ctrl+S**. **Ctrl+Shift+S** downloads a GridForge JSON backup.
- **Cell grid, address box, formula bar, and worksheet tabs** remain central. GridForge auto-saves to browser storage, but explicit XLSX/JSON exports are recommended for portable backups.
- **Insights is opt-in.** The formula inspector, Impact, What-if, Analysis panels, and per-change performance counters are hidden until the user chooses **Insights**, **View**, or **Formulas**. They have not been removed.
- **Guided & simple mode** keeps tutorial links and quick-start actions for new users. Excel familiar is the default workspace mode; switching modes does not alter workbook content.
- **Right-click** offers common cell actions; formula/functions/automations and Data Studio remain discoverable through the ribbon/search.

## Beginner walkthrough

1. Double-click `START_HERE_WINDOWS.bat` or run `python run_gridforge.py`.
2. Click a cell and enter a value. Use arrow keys, Tab, and Enter to move.
3. Select the formula bar (or type `=` into a cell) and try `=SUM(B4:B11)`.
4. Use **Ctrl+O** to open a supported XLSX/XLSM file.
5. Use **Ctrl+S** to save a new XLSX copy, or **Ctrl+Shift+S** for a GridForge-native JSON backup.
6. Use **Data** for sort/filter, **Formulas** to inspect the formula library, and **Insights** to debug a selected cell.
7. Switch to dark mode or Guided mode from the upper right without losing edits.

## Compatibility notes

The XLSM/VBA workflow still requires installed desktop Microsoft Excel for actual VBA execution. Advanced workbooks, large pivot tables, external linked data types, dynamic spilling, and every Excel formula variant are not yet fully supported. The 506-function registry includes names with limited behavior. See `docs/V10_3_EXCEL_REPAIR_AUDIT.md`.

# Reviewer walkthrough: GridForge in 3–5 minutes

Start the included app with only Python 3: on Windows double-click `START_HERE_WINDOWS.bat`, or run `python run_gridforge.py`. No npm modules are required for this demo.

1. **Grid and formula engine** — Point out visible virtualized grid and editable formula bar. Edit `B4` on sample demo; observe `D4` recalculate. Explain handwritten parser and incremental dependency graph, not an external spreadsheet evaluator.
2. **Nested formulas** — Enter `=IF(SUM(B4:B5)>100,"High","Low")` into a blank cell. Try `=SUM(A1,B2` to show a caret diagnostic. Create a direct `=A1` in A1 to show cycle handling, then Undo.
3. **Spreadsheet interactions** — Right-click a cell for the custom menu. Show copy/paste adjusting references; bold and undo; select a table and use Custom Sort with a header row.
4. **Real `.xlsx` support** — Import `examples/sample_sales.xlsx`. Switch between Sales and Summary sheets; apply a text Filter, freeze first row and column, and highlight totals with Conditional Formatting. Export XLSX and import it again.
5. **Unique tools** — Open Formula Inspector, Dependency Explorer, What-if preview, and Data Quality analysis. Explain their value for debugging rather than implying AI-based formula calculation.
6. **Proof of performance** — Open `docs/FINAL_RELEASE_EVIDENCE.md`. Show the 10,000-formula benchmark and the reported source-edit recalculation of only 26 cells, not all 10,000. Explain that this is a targeted case, not universal speed.
7. **Limitations** — Mention cross-sheet formulas and pivot tables are not supported. XLSX is a supported subset, while GridForge JSON is the safer app-specific backup.

Do not erase test data without exporting a backup. Note that browser-local data is tied to the exact host+port and browser profile. This script is a suggested manual demonstration, **not** a claim that every step was independently automated.

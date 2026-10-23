# GridForge v8 — Code review and defect audit

## Scope and method

The audit covers the v8 changes (100 Data Studio operations, UI integration and previews), regressions affecting the existing custom parser, grid and workbook history, XLSX/XLSM file interchange and optional Excel-bridge request security. The tests are automated, deterministic and run in isolated processes/browser profiles. Passing tests are evidence for the tested cases, not proof the entire codebase is free of defects.

## Confirmed defects discovered and corrected

1. **Duplicate-row formula lookup used the first matching cell's column.** If two source cells had identical raw formula text, the first occurrence could be evaluated for both cells. The row scanner now uses the actual column index for each formula. The comparison is only evaluated when the `dedupe-rows` operation is requested.
2. **Data Studio first open had no selected action.** The UI initially displayed a catalog while saying no actions matched the filter. The tool picker now selects the first option when available and provides a meaningful description immediately.
3. **Version label still displayed v6.0.** The application footer now identifies v8.0.

The XLSM serializer bug fixed earlier in v7 (duplicate XML declaration) remains covered by the XLSM readback test. No new changes were made to the VBA runtime bridge in v8.

## Automated coverage

- Exactly **100 unique commands** across **10 categories** (10 each), with an individual behavior test and expected result for every operation.
- Text, number, percent and currency conversion; fill direction and numeric sequence; square transformations and cyclic shifts; quality cleanup; formula insertion; and grouped Undo/Redo.
- Formula preservation by data cleanup, relative/absolute formula reference translation for fills, square-only preconditions and invalid bounds.
- Validation for nonnumeric input, zero denominators, negative square roots, selections larger than 12,000 cells and missing output room.
- 10,000-cell stress benchmark checks preview non-mutation, exact output, atomic Undo and unaffected outside cells.
- Browser regression checks toolbar action, modal visibility, category filtering, selection preview, apply, Undo, Redo, quick SUM insertion, formula fill, and JavaScript errors.
- Older tests continue to check 306-function parsing/evaluation, recalculation, circular errors, formatting, sorting, XLSX compatibility, XLSM package patching, and VBA bridge protocol security.

## Known limits and open verification

- **No claim of full Excel parity.** The 100 additions are workflow commands rather than new Excel formula functions. There are still 306 registered formula functions.
- The Text & case and Numeric groups intentionally skip formulas. Range rearrangements and Fill *do* move formulas; verify complex references on your own data.
- The tool preview is capped at 12,000 selected cells. Large sheets beyond the cap require working in sections.
- Some formatting, advanced Excel features, shared formulas, workbook signatures and worksheet restructuring cannot round-trip to XLSM. Original VBA binaries can be preserved, but VBA macros are not executed in the browser.
- Actual Microsoft Excel VBA execution needs a Windows computer with a trusted workbook, Excel installed and an explicitly running bridge. This environment tested protocol guards only.
- Hosted Chromium cannot verify actual localStorage persistence across reload. Verify it on Windows in the attached manual acceptance guide.
- UI and engine tests exercise important paths but do not cover every possible combination of all 306 formula functions, formats, imported workbooks and unusual file formats.
- The selected-range quick aggregate tools may overwrite cells directly below the selection; the UI asks for confirmation, but API callers should preview themselves.
- Spreadsheet data is held in browser localStorage, which has browser-specific quota limits. Export JSON backups regularly.

## Re-run from GitHub checkout

On a developer machine with Node.js and Python + Playwright/Chromium:

```powershell
npm install
npm run regression
npm run benchmark:tools
```

On a Python-only machine to smoke-test the precompiled distribution:

```powershell
python scripts/python_smoke.py
python run_gridforge.py
```

For exact results, see `docs/V8_RELEASE_EVIDENCE.md` and run logs. If any stage fails, do not declare the version fully accepted. Retain user data by exporting JSON before modifying or reinstalling the app.

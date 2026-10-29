# GridForge v11.0 — release evidence

Tests completed in hosted Linux environment:

- `npm test`: 901/901 tests passed on full suite before final release pass.
- `node --test tests/migration_v110.test.mjs`: 12/12 passed.
- `python scripts/smoke_migration_v110.py`: Chromium in-memory UI PASS: create named table, open dropdown, select exact values, evaluate `SUM(SalesData[Units])`, edit source data to expand table+chart, export/import native XLSX, verify structured formulas and preserve complex data model; no uncaught JavaScript errors.
- Example: `node scripts/build_migration_demo_v110.mjs`: generated workbook containing native table, 2 charts, multiple filters, 2-color formatting and cached `F2=3085`, `D7=600`.

Pending: Microsoft Excel COM run, Windows WebView2 UI launch, actual installer compile/install/uninstall and signing.

See `docs/CONTINUATION_GUIDE_V11_0.md` for precise reproducible commands and limitations.

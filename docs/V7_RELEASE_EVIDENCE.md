# GridForge v7.0 release evidence — 2026-10-21

## Actual test outcomes in the authoring container

- **TypeScript type checking:** PASS (`npm run check`).
- **Unit suite:** 340/340 PASS. Preserved all previous v6 functionality including 306 registered function names.
- **10,000-formula dependency benchmark:** PASS, with independent edited cells and affected dependent cells selectively recalculated. Results are environment-dependent; see `npm run benchmark` for current metrics.
- **Chromium interaction regression:** 24/24 PASS using isolated offline document injection due hosted browser navigation restrictions. This includes a real downloaded XLSM being checked for edited values and preserved VBA part bytes.
- **XLSM-specific binary round-trip:** PASS. Tested edits to text, numeric values, newly created cells and formulas; re-import of the updated XLSM; VBA binary and media part unchanged; recalculation flags set; sheet-structure changes rejected.
- **Local Excel VBA bridge protocol:** PASS (8 checks). Wrong token, disallowed non-local origin, malformed requests, invalid macro identifier, invalid XLSM and valid macro-style identifiers covered. This test does NOT execute Excel VBA.
- **Full `npm run regression`:** 6/6 stages PASS, approximately 24 seconds in the authoring container.
- **Python-only release:** smoke tested with standard-library server and compiled assets; run `python scripts/python_smoke.py` on a clean extraction for confirmation.

## Important unverified items

- **Real Microsoft Excel COM macro execution:** NOT TESTED here. The authoring platform is Linux and does not include Microsoft Excel desktop or Windows `pywin32` COM. The code supports this on Windows, but it is not an empirically verified claim until the included Windows VBA smoke test is completed.
- **Unmodified VBA binary tests use a synthetic placeholder**, not a signed real VBA project or Microsoft Excel validation. The project tests validate binary preservation, OOXML package structure, and import/reimport in GridForge. The exported workbook should be verified using Excel and a trusted real `.xlsm`.
- **Browser `localStorage` reload:** skipped by the hosted browser policy. The user's previous Windows baseline passed this test; retest locally for v7.
- Complex `.xlsm` workbook types, password protected files, signed VBA projects, XLM macros, security prompts, workbook events and externally linked data have not been exhaustively tested.

## Exact release limitations

See `docs/V7_NATIVE_XLSM_AND_VBA.md` for limits on 25 MB input size, 5,000 edited cells, shared/array formula modification, sheet-structure changes, unmerged GridForge formatting, potential invalidated document signatures, VBA Trust Center requirements and macro execution risks.

Do not label the Windows VBA runtime as fully validated until the manual acceptance test passes on actual Excel.

# GridForge v12.0.133 — Larger worksheets and compatibility audit

- Increased the sparse worksheet canvas to 100,000 rows × 1,024 columns.
- Raised per-worksheet populated/style storage and XLSX cell-record limits to 1,000,000.
- Tightened the workbook-aware `AREAS` path to validate a single cell/range reference before returning its count; union-reference counting remains unsupported.
- Reconciled the formula reference and compatibility audit to the source registry: 506 local/workbook-aware formula names plus 10 connected names. Existing `REGEXEXTRACT` supports first-match, all-match, and capture-group modes with bounded regex input.
- Documented remaining XLSX XML/import caps and the continuing gap from Excel’s full worksheet dimensions; ODS remains unsupported.
- Build/package checks only; automated tests and live UI checks were not run.

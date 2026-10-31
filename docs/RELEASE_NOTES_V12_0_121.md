# GridForge v12.0.121 — Sparse basic aggregations

`SUM`, `AVERAGE`, `MIN`, and `MAX` now evaluate local cell and range references by visiting populated cells directly, without constructing a dense matrix for the whole range. Large local references can use the configured worksheet area, while general array formulas retain the existing 20,000-cell limit.

The sparse path preserves duplicate and overlapping range arguments, evaluates spilled cells, checks referenced formula errors, and keeps empty or text-only ranges aligned with the existing aggregate behavior. Cross-sheet references and mixed literal/array arguments continue through the standard evaluator and retain the general range cap.

Verification is limited to the TypeScript check and production/installer packaging. No automated tests, desktop launch, live UI, or Excel GUI checks are performed.

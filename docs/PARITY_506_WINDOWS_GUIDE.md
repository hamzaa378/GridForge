# All 506 GridForge functions vs Microsoft Excel

## What is actually verified?

The project contains exactly **506 registered formula function names**. The 506-fixture suite invokes **each function by name** using the real GridForge Workbook/parser/evaluator. On the bundled version the local engine produced a value for **505/506** and the remaining function, `NA()`, correctly produced the expected `#N/A` error. Zero functions raised uncaught JavaScript exceptions. These results **do not establish Excel parity**.

One invocation per name is a smoke/differential test, **not comprehensive function correctness testing**. Several special functions are intentionally approximated, arrays are often wrapped with `INDEX()` to compare a scalar result, and full Excel spill behavior is not exercised. Fixture recipes reuse previously tested GridForge examples and apply a deterministic worksheet seed.

The actual **Excel comparison has NOT been executed in this build environment** because Microsoft Excel COM is available only on Windows with installed desktop Excel. There is no fabricated Excel percentage in the ZIP.

## Run on your Windows computer (with desktop Excel installed)

You already connected to Microsoft Excel **16.0** using `pywin32`. In PowerShell, `cd` into the unzipped project and activate the virtual environment where `import win32com.client` works.

```powershell
python -c "import win32com.client; print('Excel bridge ready')"
python scripts\excel_parity_506_desktop.py
python scripts\excel_parity_506_compare.py
```

The runner uses an unsaved, empty Excel workbook. It never runs VBA or opens existing files. It does not change the application's calculation mode, fixing the failure encountered in v10.1. Each function is evaluated in an isolated cleared sheet region and captured individually; a rejected formula does not abort the other cases.

**Quick smoke test (first 25 only):**

```powershell
python scripts\excel_parity_506_desktop.py --limit 25
```

The comparison script refuses to call a partial run a complete result. Run the no-limit version again afterward.

**Optional cloud testing:** `IMAGE`, `TRANSLATE` and `DETECTLANGUAGE` are deliberately marked as not run by default, because Excel may contact an online service. To opt in to Excel executing those three formulas, run:

```powershell
python scripts\excel_parity_506_desktop.py --allow-online-functions
python scripts\excel_parity_506_compare.py
```

Only opt in if your local Excel configuration and data-sharing preferences allow this. No personal workbook data is involved in the generated fixtures.

## Reports created

- `docs/PARITY_506_EXCEL_RESULTS.json`: actual Excel output, Excel version/build and per-function status
- `docs/PARITY_506_COMPARISON.json`: 506 per-function outcomes and overall audited counts
- `docs/PARITY_506_COMPARISON.md`: human-readable summary and mismatches
- `docs/PARITY_506_GRIDFORGE_RESULTS.json`: precomputed local GridForge outputs shipped with this release
- `docs/PARITY_506_COVERAGE.csv`: manifest of all registered functions, categories and scopes
- `tests/excel_parity_506_cases.json`: exact worksheet values and formula for each case

The comparator verifies a SHA-256 hash of the 506-case test manifest and rejects missing, mismatched, or partial Excel evidence. Its numeric equivalence tolerance is `abs=1e-8`, `rel=1e-7`.

## Understanding result labels

| Status | Meaning |
|---|---|
| `MATCH` | The tested result agrees with Excel within the stated tolerance |
| `DIFFERENCE` | Scalar/array-index example differs |
| `KNOWN_SEMANTIC_DIFFERENCE` | Known limited implementation differs; still counted as a difference |
| `EXCEL_UNAVAILABLE_OR_REJECTED` | Excel returned `#NAME?` or `#BLOCKED!` while GridForge returned a value; availability and syntax need review |
| `EXCEL_EXECUTION_FAILURE` | Excel COM could not evaluate/assign this case; requires investigation |
| `ONLINE_FUNCTION_NOT_RUN` | Optional cloud-connected function deliberately not invoked |
| `VOLATILE_NONDETERMINISTIC` | Random/current-time values cannot be compared for exact equality without freezing the clock/random source |

The denominator includes matches and differences, **not** skipped/unsupported functions. The report always lists every registered function. No agreement percentage represents all possible Excel inputs, formulas, versions, formatting, macro behavior, or workbooks.

## If the Excel run is slow or errors

The runner prints a checkpoint every 25 functions and writes partial evidence to `PARITY_506_EXCEL_RESULTS.json`. A partial file is never counted as a complete run. If Excel remains open after an unexpected failure, close only the test-specific instance; avoid terminating your personal Excel sessions. Run the script again after resolving the issue.

You can send back `docs/PARITY_506_COMPARISON.json` so differences can be inspected and corrected individually.

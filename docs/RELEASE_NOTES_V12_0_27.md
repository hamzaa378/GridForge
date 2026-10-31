# GridForge 12.0.27

Worksheet-local defined names can now contain non-recursive LAMBDA formulas. Use the Define Name dialog with a name such as `DoubleIt` and `=LAMBDA(x,x*2)`, then call `=DoubleIt(A1)` or pass `DoubleIt` to supported lambda-aware functions and GROUPBY/PIVOTBY aggregations. The name manager accepts regular cell/range names as before.

XLSX import/export and XLSM name updates preserve supported formula-backed names. Imported definitions are limited to direct LAMBDA expressions; recursion, workbook-global Lambda names, arbitrary formula-backed names, and some advanced name scopes remain unsupported. LAMBDA argument count, nesting, and array output limits continue to apply.

Verification scope: static TypeScript check, production build, Python syntax compilation, and Windows installer packaging with `-SkipTests`; no tests, app launch, live UI, Windows runtime, or Excel GUI checks.

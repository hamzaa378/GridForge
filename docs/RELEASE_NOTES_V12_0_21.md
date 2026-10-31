# GridForge 12.0.21

`REGEXEXTRACT` now accepts Excel's `return_mode` values: first match (`0`), all matches (`1`), and capture groups from the first match (`2`). The optional `case_sensitivity` argument supports case-sensitive (`0`) and case-insensitive (`1`) matching. Array modes spill through the existing dynamic-array engine; missing matches and invalid arguments return spreadsheet errors.

Regex processing limits input and pattern size and rejects several high-risk constructs. Excel uses PCRE2, while GridForge uses a bounded JavaScript regex subset, so not every pattern is interchangeable.

Verification: TypeScript static check and Windows packaging with `-SkipTests`. No tests, app launch, live UI, or Excel GUI checks were run.

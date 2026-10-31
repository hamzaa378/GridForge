# GridForge v12.0.106

## Counting and Data Studio improvements

- `COUNT` counts numeric range cells and ignores text, logical values, blanks, and formula errors in references. Direct numeric text and logical arguments are counted when they follow the supported locale grammar.
- `COUNTA` counts referenced error values as occupied; `COUNTBLANK` likewise does not treat an error cell as blank. Formula results that are empty text remain nonblank for `COUNTA` and blank for `COUNTBLANK`.
- Data Studio adds `Parse date`, `Parse time`, and `Parse date and time`, using the application's current locale parser. The actions are available in saved worksheet and external query pipelines, and skip formula cells.

## Verification scope

- TypeScript static check and production build only.
- Windows installer compilation uses `-SkipTests`.
- No automated tests, app launch, live UI, Windows runtime, or Excel GUI checks were run.

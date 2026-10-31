# GridForge v11.27 Feature Matrix

| Area | Status | Current implementation and limitation |
|---|---|---|
| Engineering radix conversions | Available for the documented numeric range | Adds 12 binary/octal/decimal/hex conversion functions, signed-width handling, validation and optional output padding. Locale-specific input coercion remains a broader formula compatibility concern. |
| `FLOOR` sign compatibility | Improved | Different-sign nonzero number and significance now return `#NUM!`; zero significance returns zero. This reflects the documented Excel `FLOOR` rule. |
| `FLOOR` / `CEILING` decimal multiples | Improved | Quotients close to an integer are snapped within a scale-aware floating-point tolerance before rounding. Other binary floating-point edge cases remain possible. |
| Multi-criteria formula range shapes | Improved | `SUMIFS`, `COUNTIFS`, and `AVERAGEIFS` compare full row/column dimensions, including result ranges; paired statistical/regression arrays also reject mismatched shapes. |
| Polynomial trendlines | Improved | Fits use centered/scaled X coordinates and QR factorization, avoiding raw high powers and normal-equation conditioning issues. Equation labels expose the normalized X variable. Broad Excel numerical comparisons are not available. |
| Moving-average trendlines | Improved | Internal plotted points use the observations' actual X coordinates, so irregular scatter spacing is retained. Extensions use mean spacing and repeat the endpoint average. |
| Standard Error chart bars | Improved | Series settings calculate standard error from sample spread and observation count; supported DrawingML `stdErr` metadata imports and exports. Asymmetric/custom and horizontal scatter error bars remain unsupported. |
| Excel formula/chart parity | Partial | These fixes cover selected behavior only. The full function, chart-import, chart-style, locale, date, error, precision, and workbook-object compatibility gaps remain. |

## Code-only verification

- `npm run check` — passed.
- `npm run build` — passed; regenerated `dist/`.
- `desktop/build_windows.ps1 -SkipTests` — installer/package compilation completed with Windows test phases skipped.
- No automated tests, Windows tests, app launch, live UI checks, or Excel GUI validation were run.

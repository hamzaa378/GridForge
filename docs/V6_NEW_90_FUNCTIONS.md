# GridForge v6 — 90 additional formula names

GridForge v6 registers 306 formula names, versus 216 in v5, using the project's own TypeScript engine. The new functions were selected to improve compatibility with ordinary Excel workbooks, especially contemporary **dotted names** that were previously unsupported by the tokenizer.

| Category | New names | Count |
|---|---|---:|
| Math & Engineering | `SQRTPI`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `COMBINA`, `PERMUTATIONA`, `BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`, `DELTA`, `GESTEP`, `ERF`, `ERFC`, `CEILING.MATH`, `FLOOR.MATH`, `CEILING.PRECISE`, `FLOOR.PRECISE`, `ISO.CEILING`, `GAMMA`, `GAMMALN`, `GAMMALN.PRECISE`, `SERIESSUM` | 24 |
| Statistical | `MAXIFS`, `MINIFS`, `AVERAGEA`, `MAXA`, `MINA`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `MODE.SNGL`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `PERCENTRANK.INC`, `PERCENTRANK.EXC`, `QUARTILE.INC`, `QUARTILE.EXC`, `RANK.EQ`, `RANK.AVG`, `FORECAST.LINEAR`, `SKEW.P`, `FISHER`, `FISHERINV`, `BINOM.DIST`, `POISSON.DIST`, `EXPON.DIST`, `WEIBULL.DIST`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `LOGNORM.DIST`, `LOGNORM.INV`, `CONFIDENCE.NORM`, `PHI`, `GAUSS`, `Z.TEST` | 36 |
| Text | `ENCODEURL`, `ARRAYTOTEXT`, `VALUETOTEXT`, `REGEXEXTRACT` | 4 |
| Lookup & Reference | `XMATCH`, `ROW`, `COLUMN`, `FORMULATEXT`, `INDIRECT`, `OFFSET`, `HYPERLINK`, `AREAS`, `CELL` | 9 |
| Logical | `IFNA` | 1 |
| Information | `ISFORMULA`, `NA`, `ERROR.TYPE`, `ISREF` | 4 |
| Financial | `CUMIPMT`, `CUMPRINC`, `ISPMT`, `DB`, `VDB`, `PDURATION`, `DOLLARDE`, `DOLLARFR`, `PRICEDISC`, `DISC` | 10 |
| Date & Time | `WORKDAY.INTL`, `NETWORKDAYS.INTL` | 2 |
| **Total** | | **90** |

## Examples you can enter

```excel
=STDEV.S(A1:A5)
=MAXIFS(B2:B10,A2:A10,"East")
=NORM.DIST(0,0,1,TRUE)
=IFNA(XMATCH("Product",A1:A10),0)
=INDIRECT("B3")
=FORMULATEXT(C5)
=CEILING.MATH(7.2,2)
=WORKDAY.INTL(DATE(2026,10,19),5,1)
```

The built-in gallery `examples/v6_90_function_gallery.xlsx` has one live calculated example per new function, and `examples/v6_90_function_gallery.json` retains GridForge-specific formatting and formulas.

## Compatibility notes

- These names are **registered and executable**, but their behavior is a documented subset of Excel. Counted functions include modern-name equivalents of older calculation families.
- `INDIRECT` and scalar `OFFSET` use runtime dependency tracking, so source edits invalidate dependent cells; R1C1 syntax, cross-sheet names, and OFFSET ranges larger than 1×1 are unsupported.
- `XMATCH` implements exact match, wildcard match and simple next-smaller/next-larger matches with forward/reverse searches. Binary search modes are unsupported.
- `ROW`, `COLUMN`, `FORMULATEXT`, `ISFORMULA`, `CELL` and `ERROR.TYPE` are evaluated specially so a referenced cell's result need not be evaluated prematurely.
- `HYPERLINK` returns its friendly name or URL as text; it does not open an external site. `AREAS` counts a single rectangle only.
- Probability functions use numerical approximations. The error function and inverse-normal methods provide typical spreadsheet precision, not guaranteed bit-for-bit Excel equivalence.
- `VDB`, `DB`, `PRICEDISC` and `DISC` implement simplified financial conventions rather than Excel's full set of day-count/basis and switch rules. Test real financial workbooks against Excel.
- Date functions use Excel-compatible 1900-system serial values with UTC date arithmetic; historic 1900 edge cases can differ.
- `REGEXEXTRACT` uses JavaScript regular-expression semantics. Large or pathological patterns may be slow; avoid regex from untrusted workbooks.
- Dynamic arrays, lambda expressions and spill range operators are not supported.
- Every new function has an explicit unit test in `tests/functions_v6.test.mjs`; complete signatures and samples are in [FUNCTION_REFERENCE.md](FUNCTION_REFERENCE.md).

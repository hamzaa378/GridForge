# GridForge v9 — Excel function compatibility gap (bounded comparison)

- **GridForge implemented:** 406 registered formula names, including 100 new in v9.
- **Frozen reference set:** 475 unique names = 355 historical names plus 123 curated modern names (after removing overlaps).
- **Present in comparison:** 371; **missing from comparison:** 104.
- **Historical set only:** 322 supported, 33 missing from 355 names.

**Read this carefully:** The numbers above are exact **for the included, frozen comparison file**, not the total number of Microsoft Excel functions. Microsoft changes its list with new Microsoft 365 releases, and GridForge has aliases/extra names not counted as separate Microsoft functions. Even a registered function may support only a subset of Excel behavior.

References: [Microsoft Excel functions (alphabetical)](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical), [Microsoft Excel functions by category](https://support.microsoft.com/en-us/excel/excel-functions-by-category). The historical list was bootstrapped from the publicly distributed `openpyxl` `FORMULAE` constant; a selection of newer function names was curated separately against Microsoft documentation. Neither source is a complete up-to-the-minute authoritative snapshot.

## Still missing from this comparison

| Function | Group |
| --- | --- |
| `ACCRINT` | Historical / compatibility |
| `ACCRINTM` | Historical / compatibility |
| `AGGREGATE` | Modern / newer |
| `AMORDEGRC` | Historical / compatibility |
| `AMORLINC` | Historical / compatibility |
| `ASC` | Historical / compatibility |
| `BAHTTEXT` | Historical / compatibility |
| `BESSELI` | Historical / compatibility |
| `BESSELJ` | Historical / compatibility |
| `BESSELK` | Historical / compatibility |
| `BESSELY` | Historical / compatibility |
| `BETA.DIST` | Modern / newer |
| `BETA.INV` | Modern / newer |
| `BINOM.DIST.RANGE` | Modern / newer |
| `BYCOL` | Modern / newer |
| `BYROW` | Modern / newer |
| `CHISQ.DIST` | Modern / newer |
| `CHISQ.DIST.RT` | Modern / newer |
| `CHISQ.INV` | Modern / newer |
| `CHISQ.INV.RT` | Modern / newer |
| `CHISQ.TEST` | Modern / newer |
| `CHOOSECOLS` | Modern / newer |
| `CHOOSEROWS` | Modern / newer |
| `CONFIDENCE.T` | Modern / newer |
| `COUPDAYBS` | Historical / compatibility |
| `COUPDAYS` | Historical / compatibility |
| `COUPDAYSNC` | Historical / compatibility |
| `COUPNCD` | Historical / compatibility |
| `COUPNUM` | Historical / compatibility |
| `COUPPCD` | Historical / compatibility |
| `COVARIANCE.P` | Modern / newer |
| `COVARIANCE.S` | Modern / newer |
| `CUBEKPIMEMBER` | Historical / compatibility |
| `CUBEMEMBER` | Historical / compatibility |
| `CUBEMEMBERPROPERTY` | Historical / compatibility |
| `CUBERANKEDMEMBER` | Historical / compatibility |
| `CUBESET` | Historical / compatibility |
| `CUBESETCOUNT` | Historical / compatibility |
| `CUBEVALUE` | Historical / compatibility |
| `DETECTLANGUAGE` | Modern / newer |
| `DROP` | Modern / newer |
| `ERF.PRECISE` | Modern / newer |
| `ERFC.PRECISE` | Modern / newer |
| `EXPAND` | Modern / newer |
| `F.DIST` | Modern / newer |
| `F.DIST.RT` | Modern / newer |
| `F.INV` | Modern / newer |
| `F.INV.RT` | Modern / newer |
| `F.TEST` | Modern / newer |
| `FILTER` | Modern / newer |
| `FORECAST.ETS` | Modern / newer |
| `FORECAST.ETS.CONFINT` | Modern / newer |
| `FORECAST.ETS.SEASONALITY` | Modern / newer |
| `FORECAST.ETS.STAT` | Modern / newer |
| `GAMMA.DIST` | Modern / newer |
| `GAMMA.INV` | Modern / newer |
| `GETPIVOTDATA` | Historical / compatibility |
| `GROUPBY` | Modern / newer |
| `HSTACK` | Modern / newer |
| `HYPGEOM.DIST` | Modern / newer |
| `IMAGE` | Modern / newer |
| `INFO` | Historical / compatibility |
| `INTRATE` | Historical / compatibility |
| `ISOMITTED` | Modern / newer |
| `JIS` | Historical / compatibility |
| `LAMBDA` | Modern / newer |
| `LET` | Modern / newer |
| `MAKEARRAY` | Modern / newer |
| `MAP` | Modern / newer |
| `MODE.MULT` | Modern / newer |
| `MUNIT` | Modern / newer |
| `ODDFPRICE` | Historical / compatibility |
| `ODDFYIELD` | Historical / compatibility |
| `ODDLPRICE` | Historical / compatibility |
| `ODDLYIELD` | Historical / compatibility |
| `PHONETIC` | Historical / compatibility |
| `PIVOTBY` | Modern / newer |
| `RANDARRAY` | Modern / newer |
| `REDUCE` | Modern / newer |
| `RTD` | Historical / compatibility |
| `SCAN` | Modern / newer |
| `SEQUENCE` | Modern / newer |
| `SHEET` | Modern / newer |
| `SHEETS` | Modern / newer |
| `SORT` | Modern / newer |
| `SORTBY` | Modern / newer |
| `STOCKHISTORY` | Modern / newer |
| `T.DIST` | Modern / newer |
| `T.DIST.2T` | Modern / newer |
| `T.DIST.RT` | Modern / newer |
| `T.INV` | Modern / newer |
| `T.INV.2T` | Modern / newer |
| `T.TEST` | Modern / newer |
| `TAKE` | Modern / newer |
| `TEXTSPLIT` | Modern / newer |
| `TOCOL` | Modern / newer |
| `TOROW` | Modern / newer |
| `TRANSLATE` | Modern / newer |
| `TRIMRANGE` | Modern / newer |
| `UNIQUE` | Modern / newer |
| `VSTACK` | Modern / newer |
| `WEBSERVICE` | Modern / newer |
| `WRAPCOLS` | Modern / newer |
| `WRAPROWS` | Modern / newer |

## Implemented names tracked by the comparison

ABS, ACOS, ACOSH, ADDRESS, AND, AREAS, ASIN, ASINH, ATAN, ATAN2, ATANH, AVEDEV
AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BETADIST, BETAINV, BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOM.INV, BINOMDIST
BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, CEILING, CEILING.MATH, CEILING.PRECISE, CELL, CHAR, CHIDIST, CHIINV
CHITEST, CHOOSE, CLEAN, CODE, COLUMN, COLUMNS, COMBIN, COMPLEX, CONCAT, CONCATENATE, CONFIDENCE, CONFIDENCE.NORM
CONVERT, CORREL, COS, COSH, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COVAR, CRITBINOM, CSC
CSCH, CUMIPMT, CUMPRINC, DATE, DATEDIF, DATEVALUE, DAVERAGE, DAY, DAYS, DAYS360, DB, DCOUNT
DCOUNTA, DDB, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DEGREES, DELTA, DEVSQ, DGET, DISC, DMAX
DMIN, DOLLAR, DOLLARDE, DOLLARFR, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DURATION, DVAR, DVARP, ECMA.CEILING
EDATE, EFFECT, EOMONTH, ERF, ERFC, ERROR.TYPE, EVEN, EXACT, EXP, EXPONDIST, FACT, FACTDOUBLE
FALSE, FDIST, FIND, FINDB, FINV, FISHER, FISHERINV, FIXED, FLOOR, FORECAST, FORECAST.LINEAR, FREQUENCY
FTEST, FV, FVSCHEDULE, GAMMA, GAMMADIST, GAMMAINV, GAMMALN, GCD, GEOMEAN, GESTEP, GROWTH, HARMEAN
HEX2BIN, HEX2DEC, HEX2OCT, HLOOKUP, HOUR, HYPERLINK, HYPGEOMDIST, IF, IFERROR, IMABS, IMAGINARY, IMARGUMENT
IMCONJUGATE, IMCOS, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSIN, IMSQRT
IMSUB, IMSUM, INDEX, INDIRECT, INT, INTERCEPT, IPMT, IRR, ISBLANK, ISERR, ISERROR, ISEVEN
ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISOWEEKNUM, ISPMT, ISREF, ISTEXT, KURT, LARGE
LCM, LEFT, LEFTB, LEN, LENB, LINEST, LN, LOG, LOG10, LOGEST, LOGINV, LOGNORM.DIST
LOGNORM.INV, LOGNORMDIST, LOOKUP, LOWER, MATCH, MAX, MAXA, MDETERM, MDURATION, MEDIAN, MID, MIDB
MIN, MINA, MINUTE, MINVERSE, MIRR, MMULT, MOD, MODE, MODE.SNGL, MONTH, MROUND, MULTINOMIAL
N, NA, NEGBINOMDIST, NETWORKDAYS, NETWORKDAYS.INTL, NOMINAL, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, NORMDIST, NORMINV
NORMSDIST, NORMSINV, NOT, NOW, NPER, NPV, OCT2BIN, OCT2DEC, OCT2HEX, ODD, OFFSET, OR
PEARSON, PERCENTILE, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK, PERCENTRANK.EXC, PERCENTRANK.INC, PERMUT, PI, PMT, POISSON, POISSON.DIST
POWER, PPMT, PRICE, PRICEDISC, PRICEMAT, PROB, PRODUCT, PROPER, PV, QUARTILE, QUARTILE.EXC, QUARTILE.INC
QUOTIENT, RADIANS, RAND, RANDBETWEEN, RANK, RANK.AVG, RANK.EQ, RATE, RECEIVED, REGEXEXTRACT, REGEXREPLACE, REGEXTEST
REPLACE, REPLACEB, REPT, RIGHT, RIGHTB, ROMAN, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, RSQ
SEARCH, SEARCHB, SECOND, SERIESSUM, SIGN, SIN, SINH, SKEW, SKEW.P, SLN, SLOPE, SMALL
SQRT, SQRTPI, STANDARDIZE, STDEV, STDEV.P, STDEV.S, STDEVA, STDEVP, STDEVPA, STEYX, SUBSTITUTE, SUBTOTAL
SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, SWITCH, SYD, T, TAN
TANH, TBILLEQ, TBILLPRICE, TBILLYIELD, TDIST, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN, TIME, TIMEVALUE, TINV
TODAY, TRANSPOSE, TREND, TRIM, TRIMMEAN, TRUE, TRUNC, TTEST, TYPE, UPPER, VALUE, VAR
VAR.P, VAR.S, VARA, VARP, VARPA, VDB, VLOOKUP, WEEKDAY, WEEKNUM, WEIBULL, WORKDAY, WORKDAY.INTL
XIRR, XLOOKUP, XMATCH, XNPV, YEAR, YEARFRAC, YIELD, YIELDDISC, YIELDMAT, Z.TEST, ZTEST

## Rebuild this report

1. Run `npm run build`.
2. Run `npm run functions:gaps`.
3. Review `docs/excel_comparison_catalog.json` to reproduce or extend the benchmark.

## Compatibility scope and high-priority next gaps

- Modern dynamic array formulas (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`) need genuine spill-range storage, resizing, and dependency support before they can behave like Excel.
- `LAMBDA`, `LET`, and related higher-order functions need name scoping and evaluation semantics, not simple function registrations.
- Cube functions, real-time-data providers, external data links, and certain modern web functions require services or connections unavailable in an offline spreadsheet.
- Legacy finance functions need exact Excel coupon conventions and date-system edge cases. Newly added legacy finance functions in v9 are documented as simplified subsets.
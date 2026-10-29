# GridForge v10 — Excel function compatibility gap (bounded comparison)

- **Registry recognizes:** 516 Excel-style names: 506 local-function implementations plus 10 connection-dependent names with explicit unsupported diagnostics.
- **Frozen reference set:** 475 unique names = 355 historical names plus 123 curated modern names (after removing overlaps).
- **Recognized in comparison:** 475; **unrecognized:** 0.
- **Of recognized names, 10 cannot evaluate live external data offline.** The catalog's zero-name gap is NOT functional Excel parity.
- **Historical set only:** 355 supported, 0 missing from 355 names.

**Read this carefully:** The numbers above are exact **for the included, frozen comparison file**, not the total number of Microsoft Excel functions. Microsoft changes its list with new Microsoft 365 releases, and GridForge has aliases/extra names not counted as separate Microsoft functions. Even a registered function may support only a subset of Excel behavior.

References: [Microsoft Excel functions (alphabetical)](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical), [Microsoft Excel functions by category](https://support.microsoft.com/en-us/excel/excel-functions-by-category). The historical list was bootstrapped from the publicly distributed `openpyxl` `FORMULAE` constant; a selection of newer function names was curated separately against Microsoft documentation. Neither source is a complete up-to-the-minute authoritative snapshot.

## Still missing from this comparison

| Function | Group |
| --- | --- |

## Implemented names tracked by the comparison

ABS, ACCRINT, ACCRINTM, ACOS, ACOSH, ADDRESS, AGGREGATE, AMORDEGRC, AMORLINC, AND, AREAS, ASC
ASIN, ASINH, ATAN, ATAN2, ATANH, AVEDEV, AVERAGE, AVERAGEA, AVERAGEIF, AVERAGEIFS, BAHTTEXT, BESSELI
BESSELJ, BESSELK, BESSELY, BETA.DIST, BETA.INV, BETADIST, BETAINV, BIN2DEC, BIN2HEX, BIN2OCT, BINOM.DIST, BINOM.DIST.RANGE
BINOM.INV, BINOMDIST, BITAND, BITLSHIFT, BITOR, BITRSHIFT, BITXOR, BYCOL, BYROW, CEILING, CEILING.MATH, CEILING.PRECISE
CELL, CHAR, CHIDIST, CHIINV, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, CHISQ.TEST, CHITEST, CHOOSE, CHOOSECOLS
CHOOSEROWS, CLEAN, CODE, COLUMN, COLUMNS, COMBIN, COMPLEX, CONCAT, CONCATENATE, CONFIDENCE, CONFIDENCE.NORM, CONFIDENCE.T
CONVERT, CORREL, COS, COSH, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, COUPDAYBS, COUPDAYS, COUPDAYSNC
COUPNCD, COUPNUM, COUPPCD, COVAR, COVARIANCE.P, COVARIANCE.S, CRITBINOM, CSC, CSCH, CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY
CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE, CUMIPMT, CUMPRINC, DATE, DATEDIF, DATEVALUE, DAVERAGE, DAY, DAYS
DAYS360, DB, DCOUNT, DCOUNTA, DDB, DEC2BIN, DEC2HEX, DEC2OCT, DECIMAL, DEGREES, DELTA, DETECTLANGUAGE
DEVSQ, DGET, DISC, DMAX, DMIN, DOLLAR, DOLLARDE, DOLLARFR, DPRODUCT, DROP, DSTDEV, DSTDEVP
DSUM, DURATION, DVAR, DVARP, ECMA.CEILING, EDATE, EFFECT, EOMONTH, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE
ERROR.TYPE, EVEN, EXACT, EXP, EXPAND, EXPONDIST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, FACT
FACTDOUBLE, FALSE, FDIST, FILTER, FIND, FINDB, FINV, FISHER, FISHERINV, FIXED, FLOOR, FORECAST
FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY, FORECAST.ETS.STAT, FORECAST.LINEAR, FREQUENCY, FTEST, FV, FVSCHEDULE, GAMMA, GAMMA.DIST, GAMMA.INV
GAMMADIST, GAMMAINV, GAMMALN, GCD, GEOMEAN, GESTEP, GETPIVOTDATA, GROUPBY, GROWTH, HARMEAN, HEX2BIN, HEX2DEC
HEX2OCT, HLOOKUP, HOUR, HSTACK, HYPERLINK, HYPGEOM.DIST, HYPGEOMDIST, IF, IFERROR, IMABS, IMAGE, IMAGINARY
IMARGUMENT, IMCONJUGATE, IMCOS, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSIN
IMSQRT, IMSUB, IMSUM, INDEX, INDIRECT, INFO, INT, INTERCEPT, INTRATE, IPMT, IRR, ISBLANK
ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISOMITTED, ISOWEEKNUM, ISPMT
ISREF, ISTEXT, JIS, KURT, LAMBDA, LARGE, LCM, LEFT, LEFTB, LEN, LENB, LET
LINEST, LN, LOG, LOG10, LOGEST, LOGINV, LOGNORM.DIST, LOGNORM.INV, LOGNORMDIST, LOOKUP, LOWER, MAKEARRAY
MAP, MATCH, MAX, MAXA, MDETERM, MDURATION, MEDIAN, MID, MIDB, MIN, MINA, MINUTE
MINVERSE, MIRR, MMULT, MOD, MODE, MODE.MULT, MODE.SNGL, MONTH, MROUND, MULTINOMIAL, MUNIT, N
NA, NEGBINOMDIST, NETWORKDAYS, NETWORKDAYS.INTL, NOMINAL, NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, NORMDIST, NORMINV, NORMSDIST
NORMSINV, NOT, NOW, NPER, NPV, OCT2BIN, OCT2DEC, OCT2HEX, ODD, ODDFPRICE, ODDFYIELD, ODDLPRICE
ODDLYIELD, OFFSET, OR, PEARSON, PERCENTILE, PERCENTILE.EXC, PERCENTILE.INC, PERCENTRANK, PERCENTRANK.EXC, PERCENTRANK.INC, PERMUT, PHONETIC
PI, PIVOTBY, PMT, POISSON, POISSON.DIST, POWER, PPMT, PRICE, PRICEDISC, PRICEMAT, PROB, PRODUCT
PROPER, PV, QUARTILE, QUARTILE.EXC, QUARTILE.INC, QUOTIENT, RADIANS, RAND, RANDARRAY, RANDBETWEEN, RANK, RANK.AVG
RANK.EQ, RATE, RECEIVED, REDUCE, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, REPLACE, REPLACEB, REPT, RIGHT, RIGHTB
ROMAN, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, RSQ, RTD, SCAN, SEARCH, SEARCHB, SECOND
SEQUENCE, SERIESSUM, SHEET, SHEETS, SIGN, SIN, SINH, SKEW, SKEW.P, SLN, SLOPE, SMALL
SORT, SORTBY, SQRT, SQRTPI, STANDARDIZE, STDEV, STDEV.P, STDEV.S, STDEVA, STDEVP, STDEVPA, STEYX
STOCKHISTORY, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ, SUMX2MY2, SUMX2PY2, SUMXMY2, SWITCH
SYD, T, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, T.TEST, TAKE, TAN, TANH, TBILLEQ
TBILLPRICE, TBILLYIELD, TDIST, TEXT, TEXTAFTER, TEXTBEFORE, TEXTJOIN, TEXTSPLIT, TIME, TIMEVALUE, TINV, TOCOL
TODAY, TOROW, TRANSLATE, TRANSPOSE, TREND, TRIM, TRIMMEAN, TRIMRANGE, TRUE, TRUNC, TTEST, TYPE
UNIQUE, UPPER, VALUE, VAR, VAR.P, VAR.S, VARA, VARP, VARPA, VDB, VLOOKUP, VSTACK
WEBSERVICE, WEEKDAY, WEEKNUM, WEIBULL, WORKDAY, WORKDAY.INTL, WRAPCOLS, WRAPROWS, XIRR, XLOOKUP, XMATCH, XNPV
YEAR, YEARFRAC, YIELD, YIELDDISC, YIELDMAT, Z.TEST, ZTEST

## Rebuild this report

1. Run `npm run build`.
2. Run `npm run functions:gaps`.
3. Review `docs/excel_comparison_catalog.json` to reproduce or extend the benchmark.

## Compatibility scope and high-priority next gaps

- Many modern dynamic array formulas are registered but **do not spill** into neighboring cells; use `INDEX` for individual results. Full Excel compatibility remains a separate engineering goal.
- `LAMBDA`, `LET`, and higher-order functions are limited subsets using safe labels, not full lexical closures or arbitrary Excel lambda expressions.
- Ten connected names are recognized with explicit offline errors; they require OLAP, PivotTable, RTD, licensed market data, or network services. They are NOT executed by GridForge offline.
- Legacy finance functions need exact Excel coupon conventions and date-system edge cases. Newly added legacy finance functions in v9 are documented as simplified subsets.
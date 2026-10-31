/** Excel-like navigation categories. Every registered function has exactly one category. */
import { functions, v5Categories, v6Categories, v9Categories, v10Categories } from './functions.js';
const groups = {
    'External & Connected': [],
    'Database': [],
    'Compatibility': [],
    'Other': [],
    'Financial': 'PMT PV FV NPV RATE NPER IPMT PPMT EFFECT NOMINAL SLN SYD DDB MIRR IRR XNPV XIRR FVSCHEDULE RRI'.split(' '),
    'Logical': 'IF IFERROR IFS SWITCH AND OR NOT XOR CHOOSE'.split(' '),
    'Lookup & Reference': 'VLOOKUP HLOOKUP XLOOKUP MATCH INDEX LOOKUP ROWS COLUMNS ADDRESS GETPIVOTDATA'.split(' '),
    'Date & Time': 'DATE YEAR MONTH DAY TODAY NOW TIME HOUR MINUTE SECOND DAYS EDATE EOMONTH WEEKDAY NETWORKDAYS WORKDAY DATEDIF YEARFRAC DAYS360 WEEKNUM ISOWEEKNUM DATEVALUE TIMEVALUE NETWORKDAYSINTL'.split(' '),
    'Information': 'ISBLANK ISNUMBER ISTEXT ISERR ISERROR ISNA ISEVEN ISODD ISLOGICAL ISNONTEXT TYPE'.split(' '),
    'Text': 'LEN UPPER LOWER TRIM CONCAT CONCATENATE LEFT RIGHT MID REPLACE FIND SEARCH SUBSTITUTE EXACT VALUE TEXTJOIN TEXTBEFORE TEXTAFTER CLEAN CODE CHAR UNICODE UNICHAR NUMBERVALUE PROPER REPT T N DOLLAR FIXED TEXT LENB LEFTB RIGHTB MIDB FINDB SEARCHB REPLACEB REGEXTEST REGEXREPLACE'.split(' '),
    'Statistical': 'AVERAGE MIN MAX COUNT COUNTA COUNTBLANK MEDIAN MODE LARGE SMALL STDEV STDEVP VAR VARP AVEDEV DEVSQ GEOMEAN HARMEAN SKEW KURT COVARIANCEP COVARIANCES PEARSON STEYX STANDARDIZE TRIMMEAN NORMDIST NORMSDIST NORMINV NORMSINV RANK PERCENTILE PERCENTRANK QUARTILE CORREL RSQ SLOPE INTERCEPT FORECAST'.split(' '),
    'Math & Engineering': 'SUM PRODUCT ABS ROUND ROUNDUP ROUNDDOWN SQRT POWER MOD INT CEILING FLOOR SIGN EXP LN LOG LOG10 PI SUMSQ SUMPRODUCT QUOTIENT TRUNC COMBIN FACT PERMUT GCD LCM EVEN ODD SIN COS TAN ASIN ACOS ATAN ATAN2 SINH COSH TANH RADIANS DEGREES ACOSH ASINH ATANH ACOT ACOTH COT COTH CSC CSCH SEC SECH MROUND FACTDOUBLE MULTINOMIAL ARABIC ROMAN BASE DECIMAL'.split(' ')
};
// Accept newly registered functions even if a category map has an omission.
const membership = new Map();
for (const name of ['CUBEKPIMEMBER', 'CUBEMEMBER', 'CUBEMEMBERPROPERTY', 'CUBERANKEDMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBEVALUE', 'RTD', 'STOCKHISTORY', 'WEBSERVICE'])
    groups['External & Connected'].push(name);
for (const [category, names] of Object.entries(groups))
    for (const name of names)
        if (functions[name] && !membership.has(name))
            membership.set(name, category);
for (const [category, names] of Object.entries(v10Categories))
    for (const name of names)
        if (functions[name] && !membership.has(name))
            membership.set(name, category);
for (const [category, names] of Object.entries(v9Categories))
    for (const name of names)
        if (functions[name] && !membership.has(name))
            membership.set(name, category);
for (const [category, names] of Object.entries(v6Categories))
    for (const name of names)
        if (functions[name] && !membership.has(name))
            membership.set(name, category);
for (const [category, names] of Object.entries(v5Categories))
    for (const name of names)
        if (functions[name] && !membership.has(name))
            membership.set(name, category);
export const categoryNames = ['All functions', 'Recently added', ...Object.keys(groups)];
export const functionCategory = (name) => membership.get(name) ?? 'Math & Engineering';
export const categorizedFunctions = Object.fromEntries(Object.keys(groups).map(category => [category, Object.keys(functions).filter(n => functionCategory(n) === category).sort()]));

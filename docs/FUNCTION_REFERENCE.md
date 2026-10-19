# Supported Spreadsheet Functions (56)

Every function below is implemented in `src/engine/functions.ts`; the `IF` and `IFERROR` branches are implemented in the evaluator to preserve lazy execution. This table is generated directly from the function registry, not copied from a generic Excel function list.

**Syntax:** Every formula starts with `=`. Commas separate arguments. A range uses a colon, as in `A1:B10`. References may use absolute locks, e.g. `$A$1`. Functions are case-insensitive.

| Function | What it does | Example |
| --- | --- | --- |
| **ABS** | Absolute value. | `=ABS(-8)` |
| **AND** | True if all conditions are true. | `=AND(A1>0,B1>0)` |
| **AVERAGE** | Arithmetic mean of numeric values. | `=AVERAGE(B1:B8)` |
| **AVERAGEIF** | Average numeric matches. | `=AVERAGEIF(A1:A10,">5",B1:B10)` |
| **CEILING** | Round up to a multiple. | `=CEILING(13,5)` |
| **CONCAT** | Join text or ranges. | `=CONCAT(A1," ",B1)` |
| **COUNT** | Count numeric cells. | `=COUNT(A1:A10)` |
| **COUNTA** | Count nonempty cells. | `=COUNTA(A1:A10)` |
| **COUNTBLANK** | Count blank cells in a range. | `=COUNTBLANK(A1:A10)` |
| **COUNTIF** | Count cells matching a criterion. | `=COUNTIF(A1:A10,">10")` |
| **EXACT** | Compare two strings with case sensitivity. | `=EXACT(A1,B1)` |
| **EXP** | Exponential e raised to x. | `=EXP(1)` |
| **FIND** | Find substring, case-sensitive, 1-based. | `=FIND("a","cable")` |
| **FLOOR** | Round down to a multiple. | `=FLOOR(13,5)` |
| **IF** | Choose one of two branches lazily. | `=IF(A1>10,"high","low")` |
| **IFERROR** | Fallback on a formula evaluation error. | `=IFERROR(A1/B1,0)` |
| **INDEX** | Return the value at 1-based row and column. | `=INDEX(A1:C5,2,3)` |
| **INT** | Round down to an integer. | `=INT(3.9)` |
| **ISBLANK** | Check if a cell is empty. | `=ISBLANK(A1)` |
| **ISNUMBER** | Check if a value is numeric. | `=ISNUMBER(A1)` |
| **ISTEXT** | Check if a value contains text. | `=ISTEXT(A1)` |
| **LARGE** | k-th largest numeric value. | `=LARGE(A1:A10,2)` |
| **LEFT** | Take the leftmost characters. | `=LEFT(A1,3)` |
| **LEN** | Number of characters. | `=LEN(A1)` |
| **LN** | Natural logarithm. | `=LN(10)` |
| **LOG** | Logarithm with optional base. | `=LOG(100,10)` |
| **LOWER** | Convert text to lowercase. | `=LOWER(A1)` |
| **MATCH** | Return 1-based position of an exact match. | `=MATCH("pear",A1:A10,0)` |
| **MAX** | Largest numeric value. | `=MAX(A1:A10)` |
| **MEDIAN** | Median of numeric values. | `=MEDIAN(A1:A10)` |
| **MID** | Take characters starting at a 1-based position. | `=MID(A1,2,4)` |
| **MIN** | Smallest numeric value. | `=MIN(A1:A10)` |
| **MOD** | Remainder after division. | `=MOD(17,5)` |
| **NOT** | Negate a boolean. | `=NOT(A1>10)` |
| **OR** | True if any condition is true. | `=OR(A1>0,B1>0)` |
| **PI** | The mathematical constant pi. | `=PI()` |
| **POWER** | Exponentiation. | `=POWER(2,8)` |
| **PRODUCT** | Multiply numeric values. | `=PRODUCT(A1:A3)` |
| **REPLACE** | Replace part of a text string. | `=REPLACE("hello",2,2,"a")` |
| **RIGHT** | Take the rightmost characters. | `=RIGHT(A1,3)` |
| **ROUND** | Round to a number of decimal places. | `=ROUND(3.14159,2)` |
| **ROUNDDOWN** | Round toward zero. | `=ROUNDDOWN(2.19,1)` |
| **ROUNDUP** | Round away from zero. | `=ROUNDUP(2.12,1)` |
| **SEARCH** | Find substring, case-insensitive. | `=SEARCH("A","cable")` |
| **SIGN** | Sign of a number. | `=SIGN(-4)` |
| **SMALL** | k-th smallest numeric value. | `=SMALL(A1:A10,2)` |
| **SQRT** | Square root. | `=SQRT(81)` |
| **STDEV** | Sample standard deviation. | `=STDEV(A1:A10)` |
| **SUBSTITUTE** | Replace all matching substrings. | `=SUBSTITUTE(A1,"old","new")` |
| **SUM** | Sum numeric values, ignoring text and blanks. | `=SUM(A1:A10)` |
| **SUMIF** | Sum values matching a criterion. | `=SUMIF(A1:A5,">10",B1:B5)` |
| **TRIM** | Remove extra whitespace. | `=TRIM(A1)` |
| **UPPER** | Convert text to uppercase. | `=UPPER(A1)` |
| **VALUE** | Convert numeric text to a number. | `=VALUE("4.5")` |
| **VAR** | Sample variance. | `=VAR(A1:A10)` |
| **VLOOKUP** | Find first-column match and return a column from the table. | `=VLOOKUP(E2,A2:C8,3,FALSE)` |

## Supported expression syntax

- **Numbers:** `1`, `12.5`, `.25`, `1e3`.
- **Text:** `"hello"`; double quotation marks inside text are escaped by doubling: `"She said ""yes"""`.
- **Boolean constants:** `TRUE`, `FALSE`.
- **References:** `A1`, `B12`, `AA9`, `$A$1`, `$A1`, `A$1`.
- **Ranges:** `A1:B10` within functions.
- **Arithmetic:** `+`, `-`, `*`, `/`, `^`, unary `+`/`-`.
- **Comparison:** `=`, `<>`, `!=`, `<`, `>`, `<=`, `>=`.
- **Text:** `&` joins values as strings.
- **Nesting:** `=ROUND(SUM(A1:A8)/COUNT(A1:A8),2)`.

## Semantics worth knowing

1. `IF` and `IFERROR` are evaluated lazily. Unused branches never trigger division-by-zero errors.
2. `SUM`, `AVERAGE`, and similar range aggregations ignore text and blank range elements unless otherwise stated.
3. `VLOOKUP` has **exact match by default** in this version; optional TRUE supports simple sorted ascending numeric approximate lookup. The lookup key must be in the first column of the supplied rectangular range. Missing matches return `#N/A`.
4. `MATCH` currently supports **exact matching (mode 0)**. Approximate modes are intentionally not provided.
5. `INDEX` uses 1-based row and column positions. Column defaults to 1 for one-column ranges.
6. `SUMIF` and `COUNTIF` criteria may be numbers, exact strings, or string comparisons such as `">100"` and `"<=5"`. Wildcard patterns are not implemented.
7. `STDEV` and `VAR` are **sample** statistics (denominator `n-1`) and require at least two numeric values.
8. `ROUND` follows JavaScript's `Math.round` plus a small epsilon; exact Excel rounding at negative halfway cases may differ.
9. `ISBLANK` treats empty text `""` as blank, in addition to never-edited cells.
10. Date/time functions, array-spilling, named ranges and cross-sheet references are intentionally out of scope.

## Error codes

| Error | Meaning | Common fix |
| --- | --- | --- |
| `#PARSE!` | Formula does not follow expression grammar | Read caret and suggestion, close parentheses |
| `#DIV/0!` | Division by zero | Check denominator, use `IFERROR` if appropriate |
| `#NAME?` | Unsupported or misspelled function | Check function reference |
| `#REF!` | Invalid cell or range boundary | Use a valid address inside workbook |
| `#CYCLE!` | Direct or indirect circular reference | Remove the loop |
| `#NUM!` | Invalid mathematical domain or infinite result | Verify numeric domain |
| `#VALUE!` | Incompatible types or scalar/range mismatch | Check arguments |
| `#ARG!` | Wrong function argument count/mode | Check signature |
| `#N/A` | Lookup did not find a match | Check keys or ranges |

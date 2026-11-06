# GridForge v12.0.210 — Formula compatibility corrections

`TEXTBEFORE` and `TEXTAFTER` now return empty text for empty input, honor omitted optional arguments such as `instance_num` and `if_not_found` (defaulting to `#N/A` while preserving explicit empty text), and return `#VALUE!` when the absolute `instance_num` exceeds the text length. Empty delimiters match immediately as Excel documents; `match_end` can treat the string end as a delimiter. Missing delimiters within the text still return `#N/A` or the provided `if_not_found` value.

`TEXTSPLIT` now returns `#VALUE!` for explicitly empty row or column delimiters. Omitting the column delimiter remains valid for splitting by rows. Uneven results use Excel's default `#N/A` padding even when the optional pad slot is explicitly omitted; an explicit `pad_with` argument overrides it, including an empty string. `FILTER` likewise treats an omitted `if_empty` slot as omitted (`#CALC!` when there are no results), while retaining an explicit `""` result. Regression tests cover these argument and spill-cell behaviors.

ODS export now writes `TEXTBEFORE` and `TEXTAFTER` through Calc's `COM.MICROSOFT` namespace, alongside `TEXTSPLIT`, and restores the formulas on import. LibreOffice documents these functions from version 25.8 under its Microsoft compatibility namespace ([TEXTBEFORE](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_textbefore.html), [TEXTAFTER](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_textafter.html), [TEXTSPLIT](https://help.libreoffice.org/latest/en-US/text/scalc/01/func_textsplit.html)); application-level calculation was not GUI-tested.

`NETWORKDAYS.INTL`, its `NETWORKDAYSINTL` alias, and `WORKDAY.INTL` now apply the default Saturday/Sunday weekend when the weekend argument is omitted before a holiday list. Invalid numeric weekend codes such as 0 return `#NUM!`; invalid weekend strings for `WORKDAY.INTL` return `#VALUE!`, while `NETWORKDAYS.INTL` returns 0 for `1111111`, matching Excel. See Microsoft’s [WORKDAY.INTL](https://support.microsoft.com/en-us/excel/functions/workday-intl-function) and [NETWORKDAYS.INTL](https://support.microsoft.com/en-us/excel/functions/networkdays-intl-function) references.

Dynamic-array argument defaults are corrected for `SORT`/`SORTBY`, `SEQUENCE`, `RANDARRAY`, `EXPAND`, `WRAPROWS`, and `WRAPCOLS`. Empty optional positions no longer become zero; omitted random-array dimensions default to one, equal RANDARRAY bounds return `#VALUE!`, `EXPAND` keeps its source dimensions, and wrap functions use the documented `#N/A` pad. See Microsoft references for [SORT](https://support.microsoft.com/en-us/excel/functions/sort-function), [RANDARRAY](https://support.microsoft.com/en-us/excel/functions/randarray-function), [EXPAND](https://support.microsoft.com/en-us/excel/functions/expand-function), [WRAPROWS](https://support.microsoft.com/en-us/excel/functions/wraprows-function), and [WRAPCOLS](https://support.microsoft.com/en-us/excel/functions/wrapcols-function).

All 1,123 source tests, 7 desktop code tests, the TypeScript build, both 10,000-item benchmarks, the Windows portable build, and installer compilation passed. No live UI, Excel, or LibreOffice GUI test was run.


## Continued implementation update — 2026-11-06

The bounded M engine now exposes `Table.CombineColumns`, `Table.Column`, `Table.HasColumns`, and `Table.IsEmpty`; registered list helpers `List.Distinct`, `List.IsEmpty`, `List.Contains`, and `List.PositionOf` are callable from M expressions. Index/computed column names reject case-insensitive collisions, and index columns validate start, step, optional type, overflow, and worksheet bounds.

The local DAX date subset now includes `YEAR`, `MONTH`, `DAY`, `QUARTER`, `WEEKDAY`, and `WEEKNUM`, including ISO week type 21. These accept Excel serials and ISO date strings; locale-dependent text coercion remains unsupported.

`npm test` passed all 1,157 source checks and rebuilt `dist/`. Focused Power Query and DAX suites passed 36/36. No live UI, Excel, or Calc test was run.

M pivot aggregation now rejects list-valued results instead of allowing unsupported nested data to escape into worksheet output. All 1,158 source tests pass; no live UI test was run.
`Table.DuplicateColumn` now copies bounded scalar or nested-join columns while keeping the nested lookup schema available for later expansion. Its optional type-metadata argument is explicitly unsupported.
`Table.Distinct` accepts a single column or a composite column list, and `Table.IsDistinct` checks for duplicate rows or keys without coercing text and numbers to the same value.
`Table.Combine` now rejects source-header collisions that differ only by letter case, rather than creating a malformed combined output schema.
The latest full code regression run passes all 1,160 source tests and rebuilds `dist/`. The focused Power Query and DAX suites pass 39/39. No live UI testing was performed.

The bounded M subset now supports scalar `Table.FromRows`/`Table.ToRows`, `Table.FromColumns`/`Table.ToColumns`, and `Table.FromRecords`/`Table.ToRecords` conversions with shape, missing-field, header, and cell-limit validation. The local DAX subset adds calendar `TOTALYTD`, `TOTALQTD`, and `TOTALMTD` for date-column references, expanding to the latest visible ISO/serial date in the current calendar period. Fiscal-year ends and general DAX date-table expressions are not supported. Code-only regression passes 1,165/1,165 and the frontend build succeeds; no live UI, Excel, or Calc validation was performed.

DAX `DATE(year, month, day)` now supports bounded numeric arguments, month/day rollover, the documented year adjustment for inputs below 1900, and the Excel serial-60 compatibility date. It works in scalar expressions and supported iterators. The build and full 1,165-case code-only regression suite pass; no live UI testing was performed.

`DATESYTD`, `DATESQTD`, and `DATESMTD` can now be used as bounded `CALCULATE` date filters. They expand from the latest visible ISO/serial date, clear direct date-table filters, and combine predictably with other supported `CALCULATE` filters. General table-valued DAX expressions, fiscal years, and native date hierarchies remain unsupported. The build and full 1,165-case code-only regression suite pass; no live UI, Excel, or Calc testing was performed.

Desktop connector security regressions now cover SQLite read-only authorizer restrictions, rejected extension loading and multi-statements, row limits, temporary session handles, large-integer results and formula-like text. Authenticated import validation tests reject insecure URLs and malformed Basic/Bearer inputs before network access. The code-only Python desktop suite passes 8/8; it opens no application window.

`DATESBETWEEN(dateColumn, start, end)` is now supported as a bounded `CALCULATE` date filter with inclusive serial/ISO endpoints and BLANK() bounds. It replaces direct date-table filters with the requested window and combines consistently with other supported CALCULATE filters. Node regression passes 1,165/1,165; the desktop Python suite passes 8/8. No live UI, Excel, or Calc testing was performed.

The local DAX subset adds DATESINPERIOD(dateColumn, startDate, count, DAY|MONTH|QUARTER|YEAR) as a bounded CALCULATE date filter. Inclusive ranges support negative, zero, and positive bounded integer counts; month/year shifts clamp to a valid target-month day. Node regression passes 1,165/1,165 and rebuilds dist/; no live UI, Excel, or Calc validation was performed.

The safe M date subset now includes Date.AddDays, Date.AddMonths, Date.AddQuarters, and Date.AddYears; month-based offsets clamp to the final valid target-month date and preserve serial fractions. Tests cover leap-day and month-end rollover plus invalid fractional offsets. Build and focused M tests pass; no live UI testing was performed.

The local DAX date subset now supports DATEADD(dateColumn, count, interval) as a bounded CALCULATE filter, shifting visible serial/ISO dates by DAY/MONTH/QUARTER/YEAR and retaining only shifted dates present in the date column. Regression covers filter replacement, month-end mapping and invalid fractional counts. Build and focused DAX tests pass 15/15; no live UI testing was performed.

The DAX date-filter subset adds SAMEPERIODLASTYEAR(dateColumn), a bounded CALCULATE filter that shifts visible dates back by one calendar year and matches only rows present in the model date column. Focused tests cover year shifting and argument validation. Build and DAX tests pass; no live UI testing was performed.

The bounded DAX date filters now include PARALLELPERIOD(dateColumn, count, MONTH|QUARTER|YEAR), which returns complete shifted calendar periods covering the first through last visible dates. Range scanning avoids large-argument stack risks and includes fractional serials through the last day of each selected period. Build and focused DAX regression pass; no live UI was tested.

M now includes bounded Date.From/Date.FromText conversions. Date.From normalizes an Excel 1900-system serial to a date serial; Date.FromText accepts exact ISO YYYY-MM-DD input only to avoid locale-dependent guessing. Focused M tests cover valid conversions and rejection of ambiguous text/wrong input types. No live UI testing was performed.

The bounded M list library adds List.Numbers, List.Combine, List.Transform, List.Select, List.RemoveNulls, and List.Reverse. Generated and transformed output is bounded to 12,000 scalar values; each lambdas use the _ item binding and predicates must return logical values. Focused Power Query tests pass 35/35; no live UI testing was performed.

After adding the bounded DAX date filters and additional M date/list helpers, the latest full code-only suite passes 1,168/1,168 and rebuilds dist/. Focused M tests pass 35/35, DAX tests pass 15/15, and desktop connector tests pass 8/8. No live UI or Excel/Calc application was opened.

The DAX table-input subset now supports VALUES(column) and DISTINCT(column) in COUNTROWS and X-iterator positions. They materialize a bounded one-column distinct value set from the current filter context and preserve blank as one distinct value; other general table-expression positions remain unsupported. Focused DAX tests pass 15/15; no live UI testing was performed.

DAX ALL(Table) and ALL(Column) now work as bounded COUNTROWS/X-iterator table inputs. Table form removes direct filters on that table; column form clears its direct filter and returns distinct remaining values. Relationship-propagated filters and general table expressions retain their existing subset behavior. Focused DAX tests pass 15/15; no live UI testing was performed.

The restricted DAX table-expression support now includes ADDCOLUMNS with scalar row expressions. It composes with FILTER, VALUES/DISTINCT, ALL, COUNTROWS and X iterators, validates unique output headers, and rejects nested/non-scalar results. Regression covers added arithmetic columns, filtered output and duplicate-name rejection. Focused DAX tests pass 15/15; no live UI testing was performed.

The bounded DAX table subset adds SELECTCOLUMNS(table, name, scalarExpression, ...) for projecting scalar expressions into a new virtual table consumed by COUNTROWS, FILTER or X iterators. Tests cover calculated projection values and output row counts. Build and focused DAX tests pass; no live UI testing was performed.

DAX table inputs add same-table SUMMARIZE(table, column, ...) for bounded distinct group rows under the active filter context. It accepts one to fifteen unique source-table columns and composes with COUNTROWS, FILTER and X iterators. Regression covers single and composite keys, filters, and column validation. Focused DAX tests pass 15/15; no live UI testing was performed.

To keep virtual table expressions bounded, DAX ADDCOLUMNS and SELECTCOLUMNS check output cell counts before materializing columns; SUMMARIZE stops grouping when its 1,000,000-cell budget would be exceeded. This matches the local model's total source-cell budget. Build and focused DAX tests pass 15/15; no live UI testing was performed.

The latest completed code-only regression run passes 1,168/1,168 and rebuilds dist/; Python desktop connector tests pass 8/8. Focused DAX tests pass 15/15 and focused M tests pass 35/35. No live UI, Excel, or Calc checks were performed.

The bounded M subset adds Table.Profile, returning per-source-column minimum, maximum, non-null count, null count, and type-aware distinct count. It handles empty tables and scalar mixed types within the existing column/row limits. Focused Power Query tests pass 40/40; no live UI testing was performed.

The local DAX aggregate subset adds DISTINCTCOUNTNOBLANK, counting unique nonblank values while preserving blank-on-empty-set behavior. Regression covers blank omission and empty filtered results. Focused DAX tests pass 15/15; no live UI testing was performed.

DAX MEDIAN(column) and MEDIANX(table, expression) now calculate medians over finite numeric values, skipping blanks/non-numeric results. Even-length medians use a half-sum to avoid overflow when the inputs are large. Focused model tests cover scalar and iterator results; no live UI testing was performed.

DAX now supports PRODUCT(column) and PRODUCTX(table, expression), multiplying finite numeric values while skipping blanks/non-numeric results and rejecting overflow. Focused tests cover aggregate and iterator calculations. No live UI testing was performed.

Verification after the current M/DAX additions: full 
npm test passes 1,177/1,177 and rebuilds dist/; Python connector tests pass 8/8. Focused M tests pass 40/40 and DAX tests pass 15/15. No live UI or Excel/Calc testing was performed.


The local DAX subset adds bounded CONCATENATEX(table, expression, delimiter) for source-order scalar concatenation, including blank-to-empty conversion and the 32,767-character cell limit. It rejects order-by arguments, which the local measure engine does not implement.


DAX now supports bounded numeric RANKX ranking with an explicit comparison value, ascending/descending order, and dense/skip tie behavior. Blank row results count as zero; text ranking and implicit visual-row comparison values remain unsupported.

The DAX subset includes inclusive PERCENTILE.INC and PERCENTILEX.INC over numeric values, with linear interpolation and checked 0–1 percentile bounds. Blank/non-numeric values are ignored. Exclusive variants are also implemented with rank-bound validation.

The local DAX subset adds exclusive PERCENTILE.EXC and PERCENTILEX.EXC. They use exclusive rank interpolation, reject percentile/rank bounds outside the data, and return blank for fewer than two numeric values.

The Power Query M subset adds List.Product (null-skipping, bounded finite numeric product) and numeric List.Median with null skipping, including use as scalar Table.Group/Table.Pivot aggregations. Product overflow and nonnumeric aggregate inputs are rejected.

The Power Query M subset adds List.Percentile, supporting one probability or a list of probabilities and ExcelInc/ExcelExc interpolation modes; the default is ExcelInc. Numeric bounds and ExcelExc rank eligibility are checked, and results are capped to the M list limit.


The M subset adds bounded List.Sort with explicit ascending/descending order, null-safe scalar comparisons and a stable fixed English text collation. Mixed-type lists and custom comparison functions are rejected.

M adds bounded list-window helpers List.FirstN, List.LastN, List.Skip, and List.Range; count and offset overloads require non-negative whole numbers within the existing 12,000-item limit. Predicate-based variants are not supported.

M Number.Round now accepts all five documented tie modes: ToEven, AwayFromZero, TowardZero, Up, and Down. Non-midpoint values still round to the nearest value; directed modes affect midpoint ties only.

Power Query M now also provides Number.RoundUp, Number.RoundDown, Number.RoundTowardZero, and Number.RoundAwayFromZero with optional decimal digits and null propagation. Directed functions and Number.Round use explicit bounded decimal precision.

The bounded Power Query M numeric library now includes `Number.Sign`, `Number.IntegerDivide`, `Number.Mod`, `Number.Power`, `Number.Sqrt`, `Number.Exp`, `Number.Ln`, `Number.Log`, and `Number.Log10`, with null propagation and explicit domain/finite-result checks. Modulo precision arguments and decimal arithmetic remain unsupported.

Power Query M adds bounded `Number.Factorial`, `Number.Combinations`, and `Number.Permutations`. These validate whole-number domains, preserve nulls, reject overflow, and cap iterative combinatorial work; results still use JavaScript binary floating-point precision.

The bounded query engine now supports `Table.ReplaceRows` for replacing an in-range row slice with validated scalar rows or complete records. It also supports an empty replacement to remove rows and a zero-length replacement range to insert rows.

Power Query M now includes bounded `Number.IsEven` and `Number.IsOdd` checks, with null propagation and finite numeric validation.

The M table subset adds `Table.ReverseRows` for deterministic row-order reversal while retaining the original column schema.

The M numeric subset now includes bounded trigonometric functions (`Sin`, `Cos`, `Tan`, `Asin`, `Acos`, `Atan`, `Atan2`, `Sinh`, `Cosh`, and `Tanh`) with inverse-domain, null, finite-input, and overflow checks.

M bitwise support now includes And, Or, Xor, Not, ShiftLeft and ShiftRight. This subset only accepts safe-integer operands, bounds shift counts to 0–52, and rejects results outside the safe-integer range.

The M interpreter recognizes common numeric constants E, Epsilon, PI, NaN, and positive/negative infinity, plus `Number.IsNaN`. Query results still reject non-finite worksheet cell values, and arithmetic does not yet propagate all IEEE special values.

`Number.FromText` now converts bounded invariant decimal/exponent strings and rejects unsupported culture-based parsing explicitly.

`Number.ToText` now converts numeric and special-number values to invariant general text. Its format and culture overloads remain unsupported and are rejected clearly.

Date hierarchy support adds month and weekday names (English culture subset), week-of-month with selectable `Day.*` start, and leap-year checks.

The M `#date(year, month, day)` constructor now validates Gregorian dates and encodes them as GridForge's Excel 1900-system serial for use in the existing date helpers.

Bounded calendar list generation is available through `#duration(days, hours, minutes, seconds)` and `List.Dates`, with 12,000-item and worksheet date-serial bounds.

`Date.AddWeeks` adds a whole-number week offset, using the same worksheet serial bounds as the existing date offset helpers.

`Date.ToRecord` exposes Year, Month and Day components as a record for calendar-table construction.

Power Query M adds `#datetime(year, month, day, hour, minute, second)` and `DateTime.ToRecord`. The constructor validates calendar/clock components and stores the result as a worksheet 1900-system serial with a fractional day; the record returns Year, Month, Day, Hour, Minute and Second. Time-zone-aware values, native M datetime typing, and exact decimal-second precision are not implemented.

`DateTime.FromText` now accepts exact ISO local timestamps (`YYYY-MM-DDTHH:mm:ss` with optional fractional seconds up to nine digits). It rejects offsets/timezone suffixes, locale guessing, and alternate separators. It returns the same worksheet serial representation as `#datetime`.

`DateTime.ToText` emits invariant ISO local timestamp text with millisecond fractional seconds. It does not provide culture-specific formats or timezone conversion.

`DateTime.From` accepts bounded GridForge worksheet serials and strict ISO local date/time text, with null propagation and explicit culture rejection. Numeric interpretation follows the worksheet serial system rather than Power Query's full OLE Automation semantics.

`Date.From` now parses strict local ISO date/time text through the same timezone-free serial conversion used by the new DateTime functions, avoiding host-local timezone shifts. `Date.FromText` remains ISO-date-only.

`Table.TransformColumnTypes` now uses the same timezone-free ISO conversion for date/datetime text and rejects timezone suffixes rather than invoking host-local parsing.

Bounded `DateTime.AddDays`, `DateTime.AddHours`, `DateTime.AddMinutes`, and `DateTime.AddSeconds` preserve fractional-day time-of-day while applying finite numeric offsets. Results are checked against worksheet serial bounds; daylight-saving and timezone arithmetic are not provided.

`DateTime.Date` and `DateTime.Time` split accepted local serial/date-time values into their whole-day serial and fractional-day component. Floating-point serials can introduce small fractional-second errors.

`DateTime.StartOfDay` and `DateTime.EndOfDay` return midnight and 23:59:59 for the input day. `DateTime.AddWeeks`, `AddMonths`, `AddQuarters`, and `AddYears` preserve the time fraction; fractional weeks are accepted, while month/quarter/year offsets require whole numbers and use month-end clamping.

The bounded time subset adds `#time`, `Time.From`, `Time.FromText`, `Time.ToRecord`, and `Time.Hour/Minute/Second`. Strict local time text and numeric fractions are encoded as fractional days; timezone values, culture-based parsing, and native M time types remain unsupported.

`Time.StartOfHour` and `Time.EndOfHour` return hour boundary fractions, through 00 seconds at the start and 59 seconds at the end.

`Time.ToText` adds invariant 24-hour formatting with millisecond fractional seconds; custom formats and cultures remain unsupported.

Added `Duration.TotalDays`, `Duration.TotalHours`, `Duration.TotalMinutes`, and `Duration.TotalSeconds` for bounded `#duration` values, with null propagation and finite-result validation.

Duration conversion now includes bounded `Duration.From` and strict `Duration.FromText`; `Duration.ToRecord` and component functions normalize elapsed days/hours/minutes/seconds, while total-unit functions remain available. Culture parsing and native duration tick storage remain unsupported.

`Duration.ToText` now writes normalized invariant `day.hh:mm:ss` text with fractional seconds when needed and rejects non-null deprecated format overrides.

Duration expressions now support addition/subtraction, negation, numeric scaling, and duration ordering. Serial-plus/minus-duration shifts are bounded to GridForge's worksheet date range; the evaluator still lacks M's native distinct date/time types.

Power Query M adds bounded `Table.TransformRows` projections to scalar lists or flat records, which can feed `Table.FromRecords` for row-shape changes.

Bounded Power Query M list set operations now include `List.Union`, `List.Intersect`, and `List.Difference`, with stable output order and duplicate multiplicity based on the input lists. These operations accept scalar values and reject custom equation criteria.

The bounded M list library adds `List.IsDistinct`, `List.ContainsAll`, `List.ContainsAny`, and `List.PositionOfAny`, with type-aware scalar matching and clear rejection of unsupported comparers and nested values.

The bounded M list library adds count-based `List.MaxN` and `List.MinN` for homogeneous scalar lists, with deterministic order and the existing 12,000-item guard. Predicate/comparer forms remain unsupported.

`List.RemoveMatchingItems` is now supported for bounded scalar lists. It removes all occurrences of requested values, preserves remaining input order, and rejects unsupported equation criteria.

The M `List.RemoveItems` alias now shares the bounded all-occurrences removal behavior with `List.RemoveMatchingItems`.

DAX now supports sample and population variance and standard deviation for model columns (`VAR.S`, `VAR.P`, `STDEV.S`, `STDEV.P`) and bounded iterator expressions (`VARX.S`, `VARX.P`, `STDEVX.S`, `STDEVX.P`). Blank values are skipped; sample forms require two numeric values.

M lists now support bounded `List.RemoveFirstN` and `List.RemoveLastN` count forms, including the default one-item removal. Negative counts and predicate overloads remain unsupported.

`List.PositionOf` now accepts the bounded `Occurrence.First`/`Occurrence.Last` selector, defaulting to the first match and returning -1 for no match. Comparer arguments remain unsupported.

The bounded M list subset now includes sample `List.StandardDeviation`; it validates that each item is finite numeric and requires at least two values. The scalar result is available to `Table.Pivot` as an aggregate.

M now includes bounded scalar `List.Mode` and `List.Modes`, including deterministic tied-mode selection/list ordering. Empty lists, nested values, and optional comparer expressions are rejected; scalar `List.Mode` can aggregate a Pivot table.

The bounded M subset adds `List.Positions`, `List.Repeat`, `List.Split`, and `List.Zip`, including null padding for uneven zipped lists and 12,000-cell output limits.

The bounded M table subset adds `Table.MaxN` and `Table.MinN` count forms. They sort rows by one homogeneous scalar column, preserve original ordering for ties, and reject custom criteria and predicates.

The M list subset adds `List.Alternate` with bounded skip/keep counts and initial offsets, including the form that skips an initial number of items and retains the remainder.

The bounded M list editor now supports `List.InsertRange`, `List.RemoveRange`, and `List.ReplaceRange` with zero-based indices, explicit counts, and output size validation.

M list transformations add ordered `List.ReplaceMatchingItems` pairs with scalar type-aware matching and a 12,000-comparison budget.

M adds `List.Single` and `List.SingleOrDefault` cardinality checks with explicit empty/default and multiple-item behavior. Both work as scalar Pivot aggregations when each pivot group has exactly one item.

M query transformations now include bounded `Table.InsertRows`, supporting exact-width scalar lists and complete source-column records with checked offset/size limits.


The bounded M subset adds Table.First, Table.Last, and Table.FirstValue. First/last row extraction preserves the table schema; empty-table defaults are accepted explicitly, and missing defaults are reported clearly. These are local, eager table operations and do not add query folding.

The bounded M row-window helpers now also accept each predicates for Table.FirstN, Table.LastN, Table.Skip, Table.RemoveFirstN, and Table.RemoveLastN. They stop at the first non-matching row from the appropriate edge and reject non-logical predicate results.

List row-window operations now accept bounded each conditions as well as numeric counts. First/skip operations scan from the beginning and last operations scan backward, stopping at the first non-match; predicates must return true or false.

The local M subset adds Table.MatchesAllRows and Table.MatchesAnyRows with strict logical predicates and short-circuit evaluation. Empty tables follow universal/existential logic: all rows match is true, any row matches is false.

Table.MaxN and Table.MinN now accept an each predicate after their existing scalar comparison-column sort. They return the contiguous sorted prefix that satisfies the predicate and stop at the first false row; predicate results must be logical.

List.MatchesAll and List.MatchesAny now evaluate bounded each predicates with short-circuiting and strict logical results. Empty lists return true for all-match and false for any-match.

List.MaxN and List.MinN now accept either a bounded item count or an each predicate after ordering comparable scalar values. Predicate selection stops on the first false result; mixed value types and custom comparison functions remain unsupported.

The bounded M subset adds Table.RemoveMatchingRows for removing matching complete scalar rows, using type-sensitive values so numeric 1 does not match text 1. Input records must have exactly the table's columns; optional custom equality criteria are unsupported.

Current code-only check: 1,226/1,226 JavaScript tests and 8/8 desktop packaging tests pass; the TypeScript build succeeds. No live UI testing was performed.

Table.AlternateRows now keeps an initial row prefix and applies bounded skip/take cycles, preserving the source schema and row order. Invalid, fractional, or over-limit counts are rejected.

The M subset now supports Table.Repeat, preserving row order while repeating a table's rows. Requested repetitions are checked against row and cell limits before output allocation.

Table.Range now accepts its optional nullable count. Omitting the count or passing null returns all rows from the offset; zero returns an empty range. Offsets/counts remain bounded whole numbers.

Table.SplitAt now returns the first N rows and remaining rows as two schema-preserving table values. Each piece can be consumed through existing List.First/List.Last and table functions; a raw list of tables is not loaded directly to worksheet cells.

Table.Split now partitions a bounded table into ordered page-sized table values. Empty input produces an empty page list; page size must be a positive whole number. Pages can be consumed with the existing list/table helpers.

The M subset adds Table.ToList with the default comma combiner or the existing delimiter combiner, including bounded CSV quoting. The current subset expects text-only row values and enforces the cell text length limit.

Table.FromList now converts bounded lists using its default comma splitter, an explicit delimiter splitter, or Splitter.SplitByNothing. Output columns may be supplied as names or a count; default/extra-value policies and custom splitter functions remain unsupported.

Text.ToList now converts bounded text to a list of Unicode code points, capped at 12,000 list items. It composes with the existing list transforms and does not split grapheme clusters such as combining sequences.

Record/list conversion now includes `Record.ToList` and `Record.FieldValues`, each capped at the M column limit and preserving supported record insertion order. `Table.FromList` accepts `Record.FieldValues` as its splitter and validates that all list elements are records with scalar values fitting the output schema. Nested values and typed-record metadata are outside this M subset.

The bounded record subset now includes field count/name/presence lookup, exact `Record.Field`, `Record.FieldOrDefault`, list conversions (`Record.ToList`/`Record.FieldValues`), and `Record.ToTable`. `Record.ToTable` is limited to scalar field values; record type metadata and nested values remain outside the subset.

Added bounded record construction and transformations: `Record.FromList`, `Record.FromTable`, `Record.AddField`, `Record.Combine`, `Record.SelectFields`, `Record.RemoveFields`, `Record.RenameFields`, `Record.ReorderFields`, and `Record.TransformFields`. These retain deterministic order and validate field uniqueness/missing-field behavior. Record.AddField does not implement delayed evaluation.

Added `Table.FromValue` for a scalar or bounded scalar list, including the optional `DefaultColumnName` option. The implementation rejects nested values and unsupported options.

Added `Table.ColumnCount` and `Table.ApproximateRowCount` for bounded materialized query tables. The row estimate is exact for this in-memory subset and is not an external-source estimate.

M value operators now include bounded recursive list/record equality and overloaded `&` for text concatenation, bounded list concatenation, and right-biased record merging. Record merge preserves existing field positions and appends new right-side fields; ordered and nested comparisons enforce the M nesting and collection caps.

`Table.Combine` now accepts optional explicit column names, preserving the requested order, filling absent columns with null, and omitting columns outside that schema. Header-union mode remains available when names are omitted.

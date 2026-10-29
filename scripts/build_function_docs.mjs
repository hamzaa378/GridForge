#!/usr/bin/env node
/** Rebuild a complete function reference directly from the registry.
 * Run after npm run build: node scripts/build_function_docs.mjs
 * Node built-ins only; not needed for ordinary use.
 */
import {writeFileSync} from 'node:fs';
import {functions,functionNames,v5FunctionNames,v6FunctionNames,v9FunctionNames,v10FunctionNames,connectedFunctionNames} from '../dist/engine/functions.js';
import {functionCategory} from '../dist/engine/categories.js';
import {extraFunctionNames} from '../dist/engine/functions_extra.js';
if(functionNames.length!==516 || extraFunctionNames.length!==80 || v5FunctionNames.length!==80 || v6FunctionNames.length!==90)throw Error(`Expected 516 recognized names/80 v4 extra/80 v5 extra/90 v6 extra; found ${functionNames.length}/${extraFunctionNames.length}`);
const lines=[
 '# GridForge Formula Reference — 507 Local Functions + 9 Connected Names',
 '',
 'This reference is **generated from the actual custom TypeScript function registry** by `node scripts/build_function_docs.mjs`. It includes 56 original functions, 80 added in v4, another 80 added in v5, 90 more in v6, 100 new names in v9, and 100 more registered names in v10.',
 '',
 'All formulas start with `=` and use commas to separate arguments. **507** functions have local or workbook-context implementations (some with limited behavior); the **9 External & Connected** entries are recognized with explicit errors and require services that GridForge does not provide. Do not claim full Excel parity.',
 '',
 '| Function | Category | Description | Example | Added in |',
 '| --- | --- | --- | --- | --- |',
 ...functionNames.map(name=>{
  const info=functions[name];return `| **${name}** | ${functionCategory(name)} | ${info.description.replaceAll('|','\\|')} | \`${info.example}\` | ${connectedFunctionNames.includes(name)?'v10.4 / external-only':v10FunctionNames.includes(name)?'v10':v9FunctionNames.includes(name)?'v9':v6FunctionNames.includes(name)?'v6':v5FunctionNames.includes(name)?'v5':extraFunctionNames.includes(name)?'v4':'v1–v3'} |`;
 }),
 '',
 '## External & Connected (recognized, not evaluated offline)',
 '',
 '- OLAP cube formulas require an available workbook cube connection. `GETPIVOTDATA` is available for saved GridForge Pivot Summaries with supported base aggregations and filters; its remaining limits are described in its function entry.',
 '- `RTD` needs an installed Windows COM server; `STOCKHISTORY` needs a licensed market provider; `WEBSERVICE` needs authorized network access.',
 '- These names return deterministic, helpful errors. That is not proof of behavior parity when an external connection is available.',
 '',
 '## Evaluation details',
 '',
 '- Formula grammar uses a handwritten tokenizer and Pratt parser, with nested calls, arithmetic, comparisons, references, `$` reference locks and 2D ranges.',
 '- `IF`, `IFERROR`, `IFS`, `SWITCH` and the optional `XLOOKUP` fallback evaluate unselected branches lazily.',
 '- Error codes include `#PARSE!`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NUM!`, `#N/A`, `#ARG!`, `#NAME?` and `#CYCLE!`.',
 '- Date functions use Excel’s 1900 serial system and UTC arithmetic. `TODAY()` and `NOW()` update when their formula is recalculated, not continuously.',
 '- `XLOOKUP` supports one-dimensional exact, wildcard, next-smaller and next-larger matching; forward, reverse, ascending-binary and descending-binary search; an optional not-found result; and aligned row/column return arrays. Binary searches validate sort order. Full Excel coercion, locale and dynamic-array semantics remain incomplete.',
 '- `XMATCH` supports exact, wildcard and approximate modes with forward/reverse or sorted binary search. Binary search requires data sorted in the requested direction.',
 '- `MATCH` supports exact and sorted approximate modes; the default match type is approximate ascending. It rejects incorrectly sorted approximate lookup arrays.',
 '- `VLOOKUP` and `HLOOKUP` support exact lookup and ascending approximate lookup when omitted/TRUE; approximate mode requires ascending first-row/first-column keys and returns the greatest eligible key that does not exceed the target. Exact VLOOKUP/HLOOKUP matching is case-insensitive and supports `*`, `?`, and `~` patterns.',
 '- `LOOKUP` supports approximate last-eligible matching over an ascending lookup vector; it does not implement every Excel vector/array coercion rule.',
 '- `HLOOKUP` supports exact matches; `PERCENTILE` and `QUARTILE` use inclusive interpolation; `PMT`, `PV`, `FV`, `NPV` use periodic cash flows.',
 '- Conditional functions match numbers and text with basic operators (`>`, `<`, `>=`, `<=`, `=`); common Excel `*`/`?` wildcard criteria are implemented, but array-spilling is not.',
 '- XLSX interoperability retains the custom engine and supports a subset of OOXML. Extended function names such as `_xlfn.XLOOKUP` are normalized during import.',
 '- GridForge action macros v2 support preflight-safe worksheet edits, copy/paste, fill down, sort, replace and previews. These are NOT Excel VBA.',
 '- Modern dotted function names such as `STDEV.S` and `NORM.DIST` are parsed as function names.',
 '- `SHEET()` returns the active worksheet position, `SHEET(reference)` returns the referenced worksheet position, and `SHEETS()` returns the workbook worksheet count. `SHEETS(reference)` supports a reference on one worksheet; 3D references and defined-name arguments are not implemented.',
 '- Array-producing functions such as `TRANSPOSE`, `MMULT` and `FREQUENCY` must be wrapped in `INDEX(...)` because GridForge does not support spilled results.',
 '- Legacy distribution and security pricing functions follow documented approximations and subsets; Excel coupon schedules and date conventions may differ.',
 '- `OFFSET` currently supports scalar 1×1 results; no spill arrays are implemented. Direct `HYPERLINK` formulas with safe HTTP/HTTPS/mailto targets or simple in-workbook cell references are clickable in the grid; unsupported schemes, named destinations, and nested `HYPERLINK` formulas remain text-only.',
 '- **v10 compatibility note:** Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `TEXTSPLIT` etc.) produce in-memory matrices but do **not spill**; use `INDEX` to obtain one value.',
 '- **v10 compatibility note:** `BYROW`, `BYCOL`, `MAP`, `SCAN`, `REDUCE`, `MAKEARRAY`, `LET`, and `LAMBDA` have deliberately limited, safe syntax and do not execute full Excel LAMBDA expressions.',
 '- **v10 compatibility note:** `IMAGE` returns an image label/text placeholder; `TRANSLATE` and `DETECTLANGUAGE` are restricted offline helpers; `FORECAST.ETS` is a linear trend approximation, not exponential triple smoothing.',
 '- See [v4 expansion guide](V4_FUNCTION_EXPANSION.md) and [v5 macro safety and functions](V5_FUNCTIONS_AND_MACROS.md) for details.',
 ''
];
writeFileSync(new URL('../docs/FUNCTION_REFERENCE.md',import.meta.url),lines.join('\n'),'utf8');
console.log(`Generated ${functionNames.length} functions (${v6FunctionNames.length} added in v6).`);

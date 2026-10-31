#!/usr/bin/env node
/** Rebuild a complete function reference directly from the registry.
 * Run after npm run build: node scripts/build_function_docs.mjs
 * Node built-ins only; not needed for ordinary use.
 */
import {writeFileSync} from 'node:fs';
import {functions,functionNames,v5FunctionNames,v6FunctionNames,v9FunctionNames,v10FunctionNames,connectedFunctionNames} from '../dist/engine/functions.js';
import {functionCategory} from '../dist/engine/categories.js';
import {extraFunctionNames} from '../dist/engine/functions_extra.js';
if(functionNames.length!==516 || extraFunctionNames.length!==80 || new Set(functionNames).size!==functionNames.length || connectedFunctionNames.some(name=>!functionNames.includes(name)))throw Error(`Unexpected function registry: ${functionNames.length} recognized names, ${connectedFunctionNames.length} connected names, and ${extraFunctionNames.length} v4 additions.`);
const lines=[
 `# GridForge Formula Reference — ${functionNames.length-connectedFunctionNames.length} Local Functions + ${connectedFunctionNames.length} Connected Names`,
 '',
 'This reference is **generated from the actual custom TypeScript function registry** by `node scripts/build_function_docs.mjs`. It includes the original function set and later v4, v5, v6, v9, and v10 additions. Registry groups overlap where aliases were retained, so the total is deduplicated from the live registry.',
 '',
 `All formulas start with \`=\` and support comma or semicolon function argument separators (do not mix separators within one call); semicolon-separated formulas also accept decimal commas. **${functionNames.length-connectedFunctionNames.length}** functions have local or workbook-context implementations (some with limited behavior); the **${connectedFunctionNames.length} External & Connected** entries are recognized but need workbook connections or services. Do not claim full Excel parity.`,
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
 '- Formula grammar uses a handwritten tokenizer and Pratt parser, with nested calls, arithmetic, comparisons, references, `$` reference locks, 2D ranges, comma or semicolon function-argument separators, and decimal commas in semicolon-separated formulas. Localized function names, workbook-specific locale selection, and locale-specific date parsing are incomplete.',
 '- `IF`, `IFERROR`, `IFS`, `SWITCH` and the optional `XLOOKUP` fallback evaluate unselected branches lazily.',
 '- Error literals include the standard worksheet errors and newer `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`, `#UNKNOWN!`, `#BUSY!`, and `#PYTHON!` codes; recognizing a code does not provide the external data, Python runtime or service that may have produced it.',
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
  '- Dynamic-array results spill from a formula anchor into empty cells, with blocking-cell, worksheet-boundary and cycle checks. Spill results are bounded to 20,000 cells; Excel array behavior, formula coverage and interoperability are not complete.',
 '- Legacy distribution and security pricing functions follow documented approximations and subsets; Excel coupon schedules and date conventions may differ.',
 '- `OFFSET` currently supports scalar 1×1 results; no spill arrays are implemented. Direct `HYPERLINK` formulas with safe HTTP/HTTPS/mailto targets or simple in-workbook cell references are clickable in the grid; unsupported schemes, named destinations, and nested `HYPERLINK` formulas remain text-only.',
  '- **v12 compatibility note:** Dynamic-array results spill when the function returns a matrix. Some functions still return scalars or have narrower Excel argument/coercion behavior; use `INDEX` when a single value is required.',
  '- **v12 compatibility note:** `BYROW`, `BYCOL`, `MAP`, `SCAN`, `REDUCE`, `MAKEARRAY`, `LET`, and `LAMBDA` support scoped formula expressions and worksheet-local, non-recursive named LAMBDAs. Workbook-global names, recursion, other higher-order functions, nested array results in callbacks, and Excel’s full coercion rules are not implemented.',
 '- **v10 compatibility note:** `IMAGE` returns an image label/text placeholder; `TRANSLATE` and `DETECTLANGUAGE` are restricted offline helpers; `FORECAST.ETS` is a linear trend approximation, not exponential triple smoothing.',
 '- See [v4 expansion guide](V4_FUNCTION_EXPANSION.md) and [v5 macro safety and functions](V5_FUNCTIONS_AND_MACROS.md) for details.',
 ''
];
writeFileSync(new URL('../docs/FUNCTION_REFERENCE.md',import.meta.url),lines.join('\n'),'utf8');
console.log(`Generated ${functionNames.length-connectedFunctionNames.length} local and ${connectedFunctionNames.length} connected function entries (${functionNames.length} total).`);

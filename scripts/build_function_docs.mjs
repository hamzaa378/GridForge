#!/usr/bin/env node
/** Rebuild a complete function reference directly from the registry.
 * Run after npm run build: node scripts/build_function_docs.mjs
 * Node built-ins only; not needed for ordinary use.
 */
import {writeFileSync} from 'node:fs';
import {functions,functionNames,v5FunctionNames,v6FunctionNames,v9FunctionNames,v10FunctionNames} from '../dist/engine/functions.js';
import {functionCategory} from '../dist/engine/categories.js';
import {extraFunctionNames} from '../dist/engine/functions_extra.js';
if(functionNames.length!==506 || extraFunctionNames.length!==80 || v5FunctionNames.length!==80 || v6FunctionNames.length!==90)throw Error(`Expected 506 total/80 v4 extra/80 v5 extra/90 v6 extra; found ${functionNames.length}/${extraFunctionNames.length}`);
const lines=[
 '# GridForge Supported Spreadsheet Functions — 506',
 '',
 'This reference is **generated from the actual custom TypeScript function registry** by `node scripts/build_function_docs.mjs`. It includes 56 original functions, 80 added in v4, another 80 added in v5, 90 more in v6, 100 new names in v9, and 100 more registered names in v10.',
 '',
 'All formulas start with `=` and use commas to separate arguments. Excel-style functions are **supported subsets**, not a guarantee of full Microsoft Excel compatibility.',
 '',
 '| Function | Category | Description | Example | Added in |',
 '| --- | --- | --- | --- | --- |',
 ...functionNames.map(name=>{
  const info=functions[name];return `| **${name}** | ${functionCategory(name)} | ${info.description.replaceAll('|','\\|')} | \`${info.example}\` | ${v10FunctionNames.includes(name)?'v10':v9FunctionNames.includes(name)?'v9':v6FunctionNames.includes(name)?'v6':v5FunctionNames.includes(name)?'v5':extraFunctionNames.includes(name)?'v4':'v1–v3'} |`;
 }),
 '',
 '## Evaluation details',
 '',
 '- Formula grammar uses a handwritten tokenizer and Pratt parser, with nested calls, arithmetic, comparisons, references, `$` reference locks and 2D ranges.',
 '- `IF`, `IFERROR`, `IFS`, `SWITCH` and the optional `XLOOKUP` fallback evaluate unselected branches lazily.',
 '- Error codes include `#PARSE!`, `#VALUE!`, `#REF!`, `#DIV/0!`, `#NUM!`, `#N/A`, `#ARG!`, `#NAME?` and `#CYCLE!`.',
 '- Date functions use Excel’s 1900 serial system and UTC arithmetic. `TODAY()` and `NOW()` update when their formula is recalculated, not continuously.',
 '- `XLOOKUP` supports exact one-dimensional matches and an optional not-found result; it does not support Excel’s array-spilling, fuzzy/wildcard/approximate matching or search-mode arguments.',
 '- `HLOOKUP` supports exact matches; `PERCENTILE` and `QUARTILE` use inclusive interpolation; `PMT`, `PV`, `FV`, `NPV` use periodic cash flows.',
 '- Conditional functions match numbers and text with basic operators (`>`, `<`, `>=`, `<=`, `=`); Excel wildcard criteria and array-spilling are not implemented.',
 '- XLSX interoperability retains the custom engine and supports a subset of OOXML. Extended function names such as `_xlfn.XLOOKUP` are normalized during import.',
 '- GridForge action macros v2 support preflight-safe worksheet edits, copy/paste, fill down, sort, replace and previews. These are NOT Excel VBA.',
 '- Modern dotted function names such as `STDEV.S` and `NORM.DIST` are parsed as function names.',
 '- Array-producing functions such as `TRANSPOSE`, `MMULT` and `FREQUENCY` must be wrapped in `INDEX(...)` because GridForge does not support spilled results.',
 '- Legacy distribution and security pricing functions follow documented approximations and subsets; Excel coupon schedules and date conventions may differ.',
 '- `OFFSET` currently supports scalar 1×1 results; no spill arrays are implemented. `HYPERLINK` returns text rather than creating a clickable hyperlink.',
 '- **v10 compatibility note:** Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `TEXTSPLIT` etc.) produce in-memory matrices but do **not spill**; use `INDEX` to obtain one value.',
 '- **v10 compatibility note:** `BYROW`, `BYCOL`, `MAP`, `SCAN`, `REDUCE`, `MAKEARRAY`, `LET`, and `LAMBDA` have deliberately limited, safe syntax and do not execute full Excel LAMBDA expressions.',
 '- **v10 compatibility note:** `IMAGE` returns an image label/text placeholder; `TRANSLATE` and `DETECTLANGUAGE` are restricted offline helpers; `FORECAST.ETS` is a linear trend approximation, not exponential triple smoothing.',
 '- See [v4 expansion guide](V4_FUNCTION_EXPANSION.md) and [v5 macro safety and functions](V5_FUNCTIONS_AND_MACROS.md) for details.',
 ''
];
writeFileSync(new URL('../docs/FUNCTION_REFERENCE.md',import.meta.url),lines.join('\n'),'utf8');
console.log(`Generated ${functionNames.length} functions (${v6FunctionNames.length} added in v6).`);

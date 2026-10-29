#!/usr/bin/env node
/** Reproduce the bounded Excel-function comparison from a frozen, auditable catalog.
 * IMPORTANT: This reference is intentionally not the complete Microsoft Excel catalog.
 */
import {readFileSync,writeFileSync} from 'node:fs';
import {functions,functionNames,v10FunctionNames,connectedFunctionNames} from '../dist/engine/functions.js';
const catalog=JSON.parse(readFileSync(new URL('../docs/excel_comparison_catalog.json',import.meta.url),'utf8'));
const supported=new Set(functionNames);
const legacy=catalog.legacy_names;
const newer=catalog.modern_names;
const total=[...new Set([...legacy,...newer])].sort();
const missing=total.filter(name=>!supported.has(name));
const supportedInCatalog=total.filter(name=>supported.has(name));
const lines=[
 '# GridForge v10 — Excel function compatibility gap (bounded comparison)',
 '',
 `- **Registry recognizes:** ${functionNames.length} Excel-style names: ${functionNames.length-connectedFunctionNames.length} local-function implementations plus ${connectedFunctionNames.length} connection-dependent names with explicit unsupported diagnostics.`,
 `- **Frozen reference set:** ${total.length} unique names = ${legacy.length} historical names plus ${newer.length} curated modern names (after removing overlaps).`,
 `- **Recognized in comparison:** ${supportedInCatalog.length}; **unrecognized:** ${missing.length}.`,
 `- **Of recognized names, ${connectedFunctionNames.length} cannot evaluate live external data offline.** The catalog's zero-name gap is NOT functional Excel parity.`,
 `- **Historical set only:** ${legacy.filter(x=>supported.has(x)).length} supported, ${legacy.filter(x=>!supported.has(x)).length} missing from ${legacy.length} names.`,
 '',
 '**Read this carefully:** The numbers above are exact **for the included, frozen comparison file**, not the total number of Microsoft Excel functions. Microsoft changes its list with new Microsoft 365 releases, and GridForge has aliases/extra names not counted as separate Microsoft functions. Even a registered function may support only a subset of Excel behavior.',
 '',
 'References: [Microsoft Excel functions (alphabetical)](https://support.microsoft.com/en-us/excel/excel-functions-alphabetical), [Microsoft Excel functions by category](https://support.microsoft.com/en-us/excel/excel-functions-by-category). The historical list was bootstrapped from the publicly distributed `openpyxl` `FORMULAE` constant; a selection of newer function names was curated separately against Microsoft documentation. Neither source is a complete up-to-the-minute authoritative snapshot.',
 '',
 '## Still missing from this comparison',
 '',
 '| Function | Group |', '| --- | --- |',
 ...missing.map(name=>`| \`${name}\` | ${legacy.includes(name)?'Historical / compatibility':'Modern / newer'} |`),
 '',
 '## Implemented names tracked by the comparison', '',
 ...Array.from({length:Math.ceil(supportedInCatalog.length/12)},(_,i)=>supportedInCatalog.slice(i*12,(i+1)*12).join(', ')),
 '',
 '## Rebuild this report', '',
 '1. Run `npm run build`.','2. Run `npm run functions:gaps`.','3. Review `docs/excel_comparison_catalog.json` to reproduce or extend the benchmark.',
 '',
 '## Compatibility scope and high-priority next gaps', '',
 '- Many modern dynamic array formulas are registered but **do not spill** into neighboring cells; use `INDEX` for individual results. Full Excel compatibility remains a separate engineering goal.',
 '- `LAMBDA`, `LET`, and higher-order functions are limited subsets using safe labels, not full lexical closures or arbitrary Excel lambda expressions.',
 '- Ten connected names are recognized with explicit offline errors; they require OLAP, PivotTable, RTD, licensed market data, or network services. They are NOT executed by GridForge offline.',
 '- Legacy finance functions need exact Excel coupon conventions and date-system edge cases. Newly added legacy finance functions in v9 are documented as simplified subsets.',
 ];
writeFileSync(new URL('../docs/V10_EXCEL_FUNCTION_GAP.md',import.meta.url),lines.join('\n'),'utf8');
writeFileSync(new URL('../docs/V10_GAP_SUMMARY.json',import.meta.url),JSON.stringify({appRegistered:functionNames.length,localEvaluated:functionNames.length-connectedFunctionNames.length,connectedRecognized:connectedFunctionNames,addedV10:v10FunctionNames.length,catalogCount:total.length,covered:supportedInCatalog.length,missingCount:missing.length,legacyCount:legacy.length,missingLegacy:legacy.filter(x=>!supported.has(x)).length,missing},null,2)+'\n');
console.log(`GridForge ${functionNames.length} functions; comparison catalog ${total.length}; ${supportedInCatalog.length} covered, ${missing.length} missing.`);

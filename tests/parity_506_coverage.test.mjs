import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {functions,functionNames,connectedFunctionNames} from '../dist/engine/functions.js';
const manifest=JSON.parse(fs.readFileSync(new URL('./excel_parity_506_cases.json',import.meta.url),'utf8'));
const measured=JSON.parse(fs.readFileSync(new URL('../docs/PARITY_506_GRIDFORGE_RESULTS.json',import.meta.url),'utf8'));
const frozenParityNames=functionNames.filter(name=>!connectedFunctionNames.includes(name)&&name!=='GETPIVOTDATA');
test('Frozen 506-function parity fixtures remain complete with workbook-aware GETPIVOTDATA',()=>{
 assert.equal(functionNames.length,516);
 assert.equal(manifest.cases.length,506);
 assert.equal(measured.cases.length,506);
 assert.equal(frozenParityNames.length,506);
 assert.deepEqual(new Set(manifest.cases.map(c=>c.function)),new Set(frozenParityNames));
 assert.deepEqual(new Set(measured.cases.map(c=>c.function)),new Set(frozenParityNames));
 assert.ok(functions.GETPIVOTDATA&&!connectedFunctionNames.includes('GETPIVOTDATA'));
 for(const c of manifest.cases){
  assert.ok(functions[c.function], c.function);
  assert.match(c.formula,new RegExp('(?:^|[^A-Z0-9.])'+c.function.replaceAll('.','\\.')+'\\s*\\(','i'),c.function);
  assert.ok(['scalar','array-index','limited-semantics','volatile'].includes(c.scope));
 }
});
test('All 506 GridForge cases completed; formula errors are explicit and fixture-verified',()=>{
 assert.equal(measured.metrics.registered,506);
 assert.equal(measured.metrics.executed,506);
 assert.equal(measured.metrics.exceptions,0);
 assert.equal(measured.metrics.errors,6);
 assert.equal(measured.metrics.successful,500);
 const errors=measured.cases.filter(c=>c.gridforge.type==='error');
 assert.deepEqual(errors.map(c=>[c.function,c.gridforge.value]),[['AVERAGEA','#VALUE!'],['NA','#N/A'],['ODDFPRICE','#NUM!'],['ODDFYIELD','#NUM!'],['ODDLPRICE','#NUM!'],['ODDLYIELD','#NUM!']]);
});

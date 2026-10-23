import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {functions,functionNames} from '../dist/engine/functions.js';
const manifest=JSON.parse(fs.readFileSync(new URL('./excel_parity_506_cases.json',import.meta.url),'utf8'));
const measured=JSON.parse(fs.readFileSync(new URL('../docs/PARITY_506_GRIDFORGE_RESULTS.json',import.meta.url),'utf8'));
test('Full-registry parity suite contains exactly one formula targeting every registered function',()=>{
 assert.equal(functionNames.length,506);
 assert.equal(manifest.cases.length,506);
 assert.equal(measured.cases.length,506);
 assert.deepEqual(new Set(manifest.cases.map(c=>c.function)),new Set(functionNames));
 assert.deepEqual(new Set(measured.cases.map(c=>c.function)),new Set(functionNames));
 for(const c of manifest.cases){
  assert.ok(functions[c.function], c.function);
  assert.match(c.formula,new RegExp('(?:^|[^A-Z0-9.])'+c.function.replaceAll('.','\\.')+'\\s*\\(','i'),c.function);
  assert.ok(['scalar','array-index','limited-semantics','volatile'].includes(c.scope));
 }
});
test('All 506 GridForge cases completed without a crash; NA() intentionally returns #N/A',()=>{
 assert.equal(measured.metrics.registered,506);
 assert.equal(measured.metrics.executed,506);
 assert.equal(measured.metrics.exceptions,0);
 assert.equal(measured.metrics.errors,1);
 assert.equal(measured.metrics.successful,505);
 const errors=measured.cases.filter(c=>c.gridforge.type==='error');
 assert.deepEqual(errors.map(c=>[c.function,c.gridforge.value]),[['NA','#N/A']]);
});

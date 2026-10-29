import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {colToName} from '../dist/engine/parser.js';
const suite=JSON.parse(fs.readFileSync(new URL('./excel_parity_cases.json',import.meta.url),'utf8'));
function checkResult(g,expected){
 const actual=g.error?.code??g.value;
 if(typeof expected==='number')assert.ok(typeof actual==='number'&&Math.abs(actual-expected)<=Math.max(1e-9,Math.abs(expected)*1e-9),`actual=${JSON.stringify(actual)} expected=${expected}`);
 else assert.equal(actual,expected);
}
for(const c of suite.cases){
 test(`reference probe ${c.id}`,()=>{
  const w=new Workbook(400,52);
  const edits=Object.entries(c.inputs).map(([ref,val])=>({ref,raw:val&&typeof val==='object'&&'text' in val?"'"+val.text:typeof val==='boolean'?(val?'TRUE':'FALSE'):String(val)}));
  edits.push({ref:suite.comparisonCell,raw:c.formula});
  w.applyEdits(edits,false);
  const got=w.getValue(suite.comparisonCell);
  if(c.kind==='spill'){
   assert.equal(got.error,undefined);assert.equal(got.value,c.referenceExpected);
   for(let r=0;r<c.expectedSpill.length;r++)for(let col=0;col<c.expectedSpill[r].length;col++)assert.equal(w.getValue(`${colToName(10+col)}${5+r}`).value,c.expectedSpill[r][col],`${c.id} spill cell ${r},${col}`);
  }else checkResult(got,c.referenceExpected);
 });
}
test('Excel reference suite is deterministic, bounded, and contains no external links',()=>{
 assert.equal(suite.cases.length,140);
 assert.equal(new Set(suite.cases.map(c=>c.id)).size,140);
 for(const c of suite.cases){assert.ok(c.formula.startsWith('='));assert.equal(c.formula.includes('['),false);assert.equal(c.formula.includes('http'),false);}
});

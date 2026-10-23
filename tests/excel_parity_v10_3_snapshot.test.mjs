import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
const old=JSON.parse(fs.readFileSync(new URL('../docs/EXCEL_V10_2_BASELINE_COMPARISON.json',import.meta.url))).functions;
const excel=JSON.parse(fs.readFileSync(new URL('../docs/PARITY_506_EXCEL_RESULTS.json',import.meta.url)));
const cases=JSON.parse(fs.readFileSync(new URL('./excel_parity_506_cases.json',import.meta.url)));
const byFunction=new Map(cases.cases.map(x=>[x.function,x]));
const expected=new Map(excel.cases.map(x=>[x.function,x.excel]));
function calc(x){const book=new Workbook(400,52),edits=Object.entries(x.inputs).map(([ref,v])=>({ref,raw:v&&typeof v==='object'&&'formula' in v?v.formula:v===null?'':typeof v==='boolean'?v?'TRUE':'FALSE':typeof v==='string'?"'"+v:String(v)}));edits.push({ref:cases.comparisonCell,raw:x.formula});book.applyEdits(edits,false);const v=book.getValue(cases.comparisonCell);return v.error?{type:'error',value:v.error.code}:v.value===null?{type:'blank',value:null}:{type:typeof v.value,value:v.value};}
function equals(g,e){if(g.type!==e.type)return false;if(e.type==='number')return Math.abs(g.value-e.value)<=Math.max(1e-8,1e-7*Math.abs(e.value));return g.value===e.value;}
test('v10.3: 418 previously Excel-matching functions do not regress',()=>{
 const failures=old.filter(v=>v.status==='MATCH'&&!equals(calc(byFunction.get(v.function)),expected.get(v.function))).map(v=>v.function);
 assert.deepEqual(failures,[]);
});
test('v10.3: repair at least 32 of 34 original mismatches against actual Excel 16.0 snapshot',()=>{
 const diff=old.filter(v=>v.status==='DIFFERENCE');assert.equal(diff.length,34);
 const corrected=diff.filter(v=>equals(calc(byFunction.get(v.function)),expected.get(v.function)));
 assert.ok(corrected.length>=32,`Only ${corrected.length}/34 repaired`);
 const unresolved=diff.filter(v=>!equals(calc(byFunction.get(v.function)),expected.get(v.function))).map(v=>v.function);
 assert.deepEqual(unresolved,['IMPOWER','XNPV']);
});
test('v10.3: key calculation and error semantics',()=>{
 for(const [formula,expectedValue] of [['=CHISQ.INV(0.5,3)',2.3659738843753377],['=BESSELY(1,0)',0.08825697139770806],['=PRICE(45000,45600,0.05,0.06,100,2)',98.45612123392543],['=YIELD(45000,45600,0.05,99,100,2)',0.05643959259888173],['=DISC(45000,45100,98,100)',0.0734693877551021]]){
 const b=new Workbook(400,52);b.applyEdits([{ref:'A1',raw:formula}],false);const r=b.getValue('A1');assert.equal(r.error,undefined,formula);assert.ok(Math.abs(r.value-expectedValue)<1e-8,formula);
 }
 const b=new Workbook(400,52);b.applyEdits([{ref:'A1',raw:'=AVERAGEA(TRUE,2,"x")'}],false);assert.equal(b.getValue('A1').error?.code,'#VALUE!');
});

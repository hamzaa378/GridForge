#!/usr/bin/env node
/** Evaluate the exact same fixtures we will send to desktop Excel. Uses the real Workbook parser/evaluator. */
import fs from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
const cases=JSON.parse(fs.readFileSync(new URL('../tests/excel_parity_cases.json',import.meta.url),'utf8'));
function normalize(res){if(res.error)return {type:'error',value:res.error.code}; if(res.value===null)return {type:'blank',value:null};return {type:typeof res.value,value:res.value};}
function matches(g,w){
 if(typeof w==='number')return g.type==='number' && Math.abs(g.value-w)<=Math.max(1e-9,1e-9*Math.abs(w));
 if(w===null)return g.type==='blank';
 return (typeof w==='string'&&w.startsWith('#'))?(g.type==='error'&&g.value===w):(g.value===w);
}
const results=[];
for(const c of cases.cases){
 const book=new Workbook(400,52);
 const edits=[];
 for(const [ref,value] of Object.entries(c.inputs||{})){
  const raw=value && typeof value==='object' && 'text' in value ? "'"+value.text : typeof value==='boolean' ? (value?'TRUE':'FALSE') : String(value);
  edits.push({ref,raw});
 }
 edits.push({ref:cases.comparisonCell,raw:c.formula});
 try{
  book.applyEdits(edits,false);
  const gf=normalize(book.getValue(cases.comparisonCell));
  results.push({id:c.id,category:c.category,formula:c.formula,kind:c.kind,gridforge:gf,...(c.referenceExpected===undefined?{}:{referenceExpected:c.referenceExpected,staticReferenceMatch:matches(gf,c.referenceExpected)})});
 }catch(e){results.push({id:c.id,category:c.category,formula:c.formula,kind:c.kind,gridforge:{type:'exception',value:String(e)}});}
}
const output={schema:'gridforge-excel-parity-result/1',engine:'GridForge v10 audit',excelExecuted:false,generatedAt:new Date().toISOString(),cases:results};
const target=new URL('../docs/PARITY_GRIDFORGE_RESULTS.json',import.meta.url);
fs.writeFileSync(target,JSON.stringify(output,null,2)+'\n');
const scalar=results.filter(x=>x.kind==='scalar');const matchesKnown=scalar.filter(x=>x.staticReferenceMatch===true);const differences=scalar.filter(x=>x.staticReferenceMatch===false);
console.log(`GRIDFORGE PROBES ${results.length}; static reference agrees ${matchesKnown.length}/${scalar.length}, differences ${differences.length}, known array gaps ${results.length-scalar.length}`);
for(const r of differences)console.log(`DIFFERENCE ${r.id}: ${r.formula} produced ${JSON.stringify(r.gridforge)}; reference ${JSON.stringify(r.referenceExpected)}`);
console.log(`Wrote ${target.pathname}`);

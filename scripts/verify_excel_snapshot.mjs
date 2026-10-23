// Reference data is the user's actual Excel 16.0 run. Do not edit reference values.
import fs from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
const fixtures=JSON.parse(fs.readFileSync(new URL('../tests/excel_parity_506_cases.json',import.meta.url)));
const observed=JSON.parse(fs.readFileSync(new URL('../docs/PARITY_506_EXCEL_RESULTS.json',import.meta.url)));
const byId=new Map(observed.cases.map(c=>[c.id,c]));
const results=[];
function equal(g,e){if(g.type!==e.type)return false;if(g.type==='number')return Math.abs(g.value-e.value)<=Math.max(1e-8,1e-7*Math.abs(e.value));return g.value===e.value;}
for(const c of fixtures.cases){let w=new Workbook(400,52);const edits=Object.entries(c.inputs).map(([ref,v])=>({ref,raw:v&&typeof v==='object'&&'formula' in v?v.formula:v===null?'':typeof v==='boolean'?(v?'TRUE':'FALSE'):typeof v==='string'?"'"+v:String(v)}));edits.push({ref:fixtures.comparisonCell,raw:c.formula});
 let g;try{w.applyEdits(edits,false);const r=w.getValue(fixtures.comparisonCell);g=r.error?{type:'error',value:r.error.code}:r.value===null?{type:'blank',value:null}:{type:typeof r.value,value:r.value};}catch(e){g={type:'exception',value:String(e)};}
 const excel=byId.get(c.id)?.excel;results.push({id:c.id,function:c.function,formula:c.formula,gridforge:g,excel,match:excel?equal(g,excel):null});
}
const original=JSON.parse(fs.readFileSync(new URL('../docs/EXCEL_V10_2_BASELINE_COMPARISON.json',import.meta.url))).functions;
const beforeMatches=original.filter(x=>x.status==='MATCH'),beforeDiff=original.filter(x=>x.status==='DIFFERENCE'),known=original.filter(x=>x.status==='KNOWN_SEMANTIC_DIFFERENCE');
const current=new Map(results.map(x=>[x.function,x]));
const regressions=beforeMatches.filter(x=>!current.get(x.function)?.match).map(x=>x.function);
const fixed=beforeDiff.filter(x=>current.get(x.function)?.match).map(x=>x.function);
const pending=beforeDiff.filter(x=>!current.get(x.function)?.match).map(x=>x.function);
const summary={baselineMatches:beforeMatches.length,baselineDifferences:beforeDiff.length,newlyMatching:fixed.length,currentMatches:beforeMatches.length-regressions.length+fixed.length,comparable:beforeMatches.length+beforeDiff.length+known.length,knownSemantics:known.map(x=>x.function),regressions,unresolved:pending,source:'User-supplied actual Excel 16.0 results; new GridForge 10.3 re-evaluated locally. No new Excel session.'};
fs.writeFileSync(new URL('../docs/EXCEL_SNAPSHOT_V10_3.json',import.meta.url),JSON.stringify({summary,cases:results},null,2)+'\n');
console.log(JSON.stringify(summary,null,2));
if(regressions.length>0||fixed.length<32)process.exitCode=1;

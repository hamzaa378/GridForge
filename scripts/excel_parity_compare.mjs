#!/usr/bin/env node
/** Compare exact result types, numerical tolerances, and Excel error values. Never fabricate Excel run data. */
import fs from 'node:fs';
const root=new URL('../',import.meta.url);
const suite=JSON.parse(fs.readFileSync(new URL('tests/excel_parity_cases.json',root)));
const gf=JSON.parse(fs.readFileSync(new URL('docs/PARITY_GRIDFORGE_RESULTS.json',root)));
const path=new URL('docs/PARITY_EXCEL_RESULTS.json',root);
if(!fs.existsSync(path)){
 console.log('ACTUAL EXCEL PARITY: NOT YET MEASURED. Run scripts/excel_parity_desktop.py on Windows with desktop Excel first.');
 process.exit(2);
}
const xl=JSON.parse(fs.readFileSync(path));
if(xl.excelExecuted!==true||xl.engine!=='Microsoft Excel desktop')throw Error('Excel evidence is not an actual desktop Excel run.');
const b=new Map(gf.cases.map(c=>[c.id,c]));const e=new Map(xl.cases.map(c=>[c.id,c]));
const evaluated=[];
const errorAliases=new Map([['#N/A','#N/A'],['#N/A!','#N/A']]);
function comparable(a,z){
 if(!a||!z)return false;
 if(a.type==='number'&&z.type==='number')return Math.abs(a.value-z.value)<=Math.max(1e-8,Math.max(Math.abs(a.value),Math.abs(z.value))*1e-8);
 if(a.type==='error'&&z.type==='error')return (errorAliases.get(a.value)||a.value)===(errorAliases.get(z.value)||z.value);
 return a.type===z.type&&a.value===z.value;
}
for(const c of suite.cases){
 const gr=b.get(c.id)?.gridforge,ex=e.get(c.id)?.excel;
 let status='MATCH';
 if(!gr||!ex)status='MISSING_EVIDENCE';
 else if(c.kind==='expected-gap')status='KNOWN_LIMITATION';
 else if(ex.type==='error'&&['#NAME?','#BLOCKED!'].includes(ex.value))status='EXCEL_VERSION_OR_SECURITY_UNSUPPORTED';
 else if(!comparable(gr,ex))status='DIFFERENCE';
 evaluated.push({id:c.id,formula:c.formula,category:c.category,status,gridforge:gr,excel:ex});
}
const statuses=['MATCH','DIFFERENCE','KNOWN_LIMITATION','EXCEL_VERSION_OR_SECURITY_UNSUPPORTED','MISSING_EVIDENCE'];
const totals=Object.fromEntries(statuses.map(s=>[s,evaluated.filter(x=>x.status===s).length]));
const denominator=totals.MATCH+totals.DIFFERENCE;
const summary={excelExecuted:true,excelVersion:xl.excelVersion,excelBuild:xl.excelBuild,total:evaluated.length,totals,comparedScalarCases:denominator,scalarAgreementPercent:denominator?+(100*totals.MATCH/denominator).toFixed(2):null};
const out={schema:'gridforge-excel-parity-comparison/1',summary,cases:evaluated};
fs.writeFileSync(new URL('docs/PARITY_COMPARISON.json',root),JSON.stringify(out,null,2)+'\n');
const lines=['# Measured GridForge vs Microsoft Excel behavior','',`Actual Excel ${xl.excelVersion} (build ${xl.excelBuild}).`,`**Comparable scalar agreement: ${totals.MATCH}/${denominator} (${summary.scalarAgreementPercent}%).**`,'',...statuses.map(s=>`- ${s}: ${totals[s]}`),'','## Differences','',...evaluated.filter(x=>x.status==='DIFFERENCE').map(x=>`- **${x.id}** \`${x.formula}\`: GridForge ${JSON.stringify(x.gridforge)}, Excel ${JSON.stringify(x.excel)}`),'','## Scope','', 'This percentage applies only to the deterministic scalar fixtures provided; it is **not** the percentage of Microsoft Excel fully implemented. Dynamic arrays, workbook object model, VBA, charts, pivot tables, external links and formatting are not included in the scalar agreement denominator.'];
fs.writeFileSync(new URL('docs/PARITY_COMPARISON.md',root),lines.join('\n')+'\n');
console.log(JSON.stringify(summary,null,2));
process.exit(totals.DIFFERENCE>0?1:0);

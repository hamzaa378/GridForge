#!/usr/bin/env node
/** Reproducible, evaluated gallery for exactly the 90 functions introduced in v6. */
import {readFileSync,writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {functionCategory} from '../dist/engine/categories.js';
import {v6FunctionNames} from '../dist/engine/functions.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
const file=readFileSync(new URL('../tests/functions_v6.test.mjs',import.meta.url),'utf8');
const cases=[...file.matchAll(/\['([A-Z0-9.]+)','(=[^']*)',/g)].map(m=>({name:m[1],formula:m[2]}));
if(cases.length!==90||new Set(cases.map(x=>x.name)).size!==90||cases.some(x=>!v6FunctionNames.includes(x.name)))throw Error(`Expected exactly 90 unique v6 function examples; found ${cases.length}.`);
const sheet=new Workbook(400,52),edits=[];
for(let i=1;i<=5;i++){edits.push({ref:`A${i}`,raw:String(i)},{ref:`B${i}`,raw:String(i+1)},{ref:`D${i}`,raw:i%2?'East':'West'});}
edits.push({ref:'E2',raw:'=SUM(A1:A2)'},{ref:'F1',raw:'Function'},{ref:'G1',raw:'Formula source'},{ref:'H1',raw:'Live calculation'},{ref:'I1',raw:'Category'});
for(const [i,c]of cases.entries()){
 const row=i+2;edits.push({ref:`F${row}`,raw:c.name},{ref:`G${row}`,raw:"'"+c.formula},{ref:`H${row}`,raw:c.formula},{ref:`I${row}`,raw:functionCategory(c.name)});
}
sheet.applyEdits(edits,false);
for(let i=0;i<cases.length;i++){const name=cases[i].name,out=sheet.getValue(`H${i+2}`);if(out.error&&name!=='NA')throw Error(`v6 gallery ${name}: ${out.error.code} ${out.error.message}`);}
sheet.patchStyles(['F1','G1','H1','I1'],{bold:true,background:'#23568c',color:'#ffffff'});
const doc=new SpreadsheetDocument(sheet);doc.renameSheet(doc.activeId,'90 Functions v6');doc.active.tools.freezeRows=1;
writeFileSync(new URL('../examples/v6_90_function_gallery.json',import.meta.url),JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(new URL('../examples/v6_90_function_gallery.xlsx',import.meta.url),exportXlsx(doc));
console.log('Generated 90-function gallery JSON + XLSX.');

#!/usr/bin/env node
import {writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {functionCategory} from '../dist/engine/categories.js';
import {v9FunctionNames} from '../dist/engine/functions.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {formulas} from '../tests/v9_examples.mjs';
if(Object.keys(formulas).length!==100||v9FunctionNames.some(name=>!formulas[name]))throw Error('Gallery requires one valid formula for each of the 100 new functions.');
const book=new Workbook(400,52),edits=[];
for(let r=1;r<=5;r++){edits.push({ref:`A${r}`,raw:String(r)},{ref:`B${r}`,raw:String(r+1)},{ref:`J${r}`,raw:'0.2'});}
for(const [ref,raw] of Object.entries({H1:'1',I1:'2',H2:'3',I2:'4',D1:'Region',E1:'Sales',D2:'East',E2:'20',D3:'West',E3:'30',D4:'East',E4:'40',G1:'Region',G2:'East',K1:'Function',L1:'Live result',M1:'Original formula',N1:'Category'}))edits.push({ref,raw});
for(const [idx,name]of v9FunctionNames.entries()){
 const row=idx+2;edits.push({ref:`K${row}`,raw:name},{ref:`L${row}`,raw:formulas[name]},{ref:`M${row}`,raw:"'"+formulas[name]},{ref:`N${row}`,raw:functionCategory(name)});
}
book.applyEdits(edits,false);
for(let idx=0;idx<v9FunctionNames.length;idx++){
 const result=book.getValue(`L${idx+2}`);if(result.error)throw Error(`${v9FunctionNames[idx]} (${formulas[v9FunctionNames[idx]]}): ${result.error.message}`);
}
book.patchStyles(['K1','L1','M1','N1'],{bold:true,background:'#285b73',color:'#ffffff'});
const doc=new SpreadsheetDocument(book);doc.renameSheet(doc.activeId,'100 Functions v9');doc.active.tools.freezeRows=1;
writeFileSync(new URL('../examples/v9_100_function_gallery.json',import.meta.url),JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(new URL('../examples/v9_100_function_gallery.xlsx',import.meta.url),exportXlsx(doc));
console.log('Generated 100-function v9 gallery as JSON and XLSX; all 100 results evaluated.');

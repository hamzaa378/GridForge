#!/usr/bin/env node
/** Build a reproducible live gallery for the 80 functions added in v5. */
import {readFileSync,writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {shiftFormula} from '../dist/engine/parser.js';
import {functionCategory} from '../dist/engine/categories.js';
import {v5FunctionNames} from '../dist/engine/functions.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
const source=readFileSync(new URL('../tests/functions_v5.test.mjs',import.meta.url),'utf8');
const caseSource=source.slice(source.indexOf('const cases=['),source.indexOf('\n];',source.indexOf('const cases=[')));
const cases=[...caseSource.matchAll(/\['([A-Z0-9_]+)',\s*'(=[^']*)',/g)].map(match=>({name:match[1],formula:match[2]}));
const seen=new Set(cases.map(c=>c.name));
if(cases.length!==80||seen.size!==80||[...seen].some(name=>!v5FunctionNames.includes(name)))throw Error('Gallery must contain exactly the 80 tested new function names.');
const book=new Workbook(400,52);
const examples={AA1:'1',AA2:'2',AA3:'3',AA4:'4',AA5:'5',AB1:'3',AB2:'5',AB3:'7',AB4:'9',AB5:'11',
 AF1:'-1000',AF2:'300',AF3:'400',AF4:'500',AF5:'200',AG1:'=DATE(2020,1,1)',AG2:'=DATE(2021,1,1)',AG3:'=DATE(2022,1,1)',AG4:'=DATE(2023,1,1)',AG5:'=DATE(2024,1,1)'};
const edits=[{ref:'A1',raw:'Function'},{ref:'B1',raw:'Formula (original source)'},{ref:'C1',raw:'Live value'},{ref:'D1',raw:'Excel-style group'},...Object.entries(examples).map(([ref,raw])=>({ref,raw}))];
for(const [i,c] of cases.entries()){const row=i+2;edits.push({ref:`A${row}`,raw:c.name},{ref:`B${row}`,raw:"'"+c.formula},{ref:`C${row}`,raw:shiftFormula(c.formula,0,26)},{ref:`D${row}`,raw:functionCategory(c.name)});}
book.applyEdits(edits,false);
for(const [i,c]of cases.entries()){const v=book.getValue(`C${i+2}`);if(v.error)throw Error(`${c.name}: ${v.error.code} ${v.error.message}`);}
book.patchStyles(['A1','B1','C1','D1'],{bold:true,background:'#1e487d',color:'#ffffff'});
const doc=new SpreadsheetDocument(book);doc.renameSheet(doc.activeId,'80 New Functions');doc.active.tools.freezeRows=1;
writeFileSync(new URL('../examples/v5_80_function_gallery.json',import.meta.url),JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(new URL('../examples/v5_80_function_gallery.xlsx',import.meta.url),exportXlsx(doc));
console.log('Built v5 live gallery: 80 evaluated formulas in JSON and XLSX.');

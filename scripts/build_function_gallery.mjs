#!/usr/bin/env node
/** Creates reproducible, importable v4 formula demo files from the tested cases. */
import {readFileSync,writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {shiftFormula} from '../dist/engine/parser.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
const tests=readFileSync(new URL('../tests/functions_v4.test.mjs',import.meta.url),'utf8');
const cases=[...tests.matchAll(/^\s*\['([A-Z0-9_]+)','(=[^']*)',/gm)].map(match=>({name:match[1],formula:match[2]}));
if(cases.length!==80 || new Set(cases.map(x=>x.name)).size!==80)throw Error('Expected 80 distinct example functions.');
const book=new Workbook(400,52);
const source={
 AA1:'1',AA2:'2',AA3:'3',AA4:'4',AA5:'5',AB1:'3',AB2:'5',AB3:'7',AB4:'9',AB5:'11',
 AC1:'North',AC2:'North',AC3:'South',AC4:'North',AC5:'South',AD1:'k1',AD2:'k2',AD3:'k3',
 AE1:'10',AE2:'20',AE3:'30',AH1:'Q1',AI1:'Q2',AJ1:'Q3',AH2:'22',AI2:'33',AJ2:'44'
};
const edits=[
 {ref:'A1',raw:'Function'},{ref:'B1',raw:'Formula source'},{ref:'C1',raw:'Live GridForge result'},{ref:'D1',raw:'Category'},
 ...Object.entries(source).map(([ref,raw])=>({ref,raw})),
];
const categories=[['Math',24],['Statistics',16],['Text',12],['Dates and time',16],['Logic and lookup',8],['Finance',4]];
let i=0;
for(const [category,count] of categories)for(let j=0;j<count;j++){
 const c=cases[i++],row=i+1;
 // Internal input ranges live in AA:AJ so that the gallery table itself never
 // changes the referenced values. Relative coordinates are shifted 26 columns.
 const shifted=shiftFormula(c.formula,0,26);
 edits.push({ref:`A${row}`,raw:c.name},{ref:`B${row}`,raw:"'"+c.formula},{ref:`C${row}`,raw:shifted},{ref:`D${row}`,raw:category});
}
book.applyEdits(edits,false);
const failures=cases.flatMap((c,i)=>{const value=book.getValue(`C${i+2}`);return value.error?[`${c.name}: ${value.error.code} ${value.error.message}`]:[];});
if(failures.length)throw Error(`Gallery has invalid formulas:\n${failures.join('\n')}`);
book.patchStyles(['A1','B1','C1','D1'],{bold:true,background:'#1e487d',color:'#ffffff'});
const doc=new SpreadsheetDocument(book);doc.renameSheet(doc.activeId,'136 Functions');doc.active.tools.freezeRows=1;
const path=new URL('../examples/',import.meta.url);
writeFileSync(new URL('v4_80_function_gallery.json',path),JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(new URL('v4_80_function_gallery.xlsx',path),exportXlsx(doc));
console.log(`Built both JSON and XLSX galleries with ${cases.length} evaluated formulas.`);

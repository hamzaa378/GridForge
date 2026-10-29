import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {validateMacro,runActionMacro} from '../dist/engine/macros.js';
const make=()=>{const d=new SpreadsheetDocument();d.active.book.applyEdits([{ref:'A1',raw:'5'},{ref:'B1',raw:'=A1*2'},{ref:'C1',raw:'A'}]);return d;};
const base={version:3,name:'Offline macro',sheetName:'Sheet 1'};
test('v3 number series applies in row-major order and supports one-step undo',()=>{
 const d=make(),m={...base,steps:[{kind:'sequence',range:'D2:E3',start:10,increment:5}]};
 const preview=runActionMacro(d,m,{dryRun:true});assert.equal(preview.dryRun,true);assert.equal(d.active.book.getRaw('D2'),'');
 const report=runActionMacro(d,m);assert.equal(report.values,4);
 assert.deepEqual(['D2','E2','D3','E3'].map(r=>d.active.book.getRaw(r)),['10','15','20','25']);
 assert.equal(d.active.book.undo(),true);assert.equal(d.active.book.getRaw('D2'),'');
});
test('v3 evaluates formula result in ifCell, then and else branches',()=>{
 const d=make();runActionMacro(d,{...base,steps:[{kind:'ifCell',ref:'B1',operator:'greaterThan',value:'9',thenSteps:[{kind:'cells',edits:[{ref:'D1',raw:'HIGH'}]}],elseSteps:[{kind:'cells',edits:[{ref:'D1',raw:'LOW'}]}]}]});
 assert.equal(d.active.book.getRaw('D1'),'HIGH');d.active.book.applyEdits([{ref:'A1',raw:'2'}]);
 runActionMacro(d,{...base,steps:[{kind:'ifCell',ref:'B1',operator:'greaterThan',value:'9',thenSteps:[{kind:'cells',edits:[{ref:'D1',raw:'HIGH'}]}],elseSteps:[{kind:'cells',edits:[{ref:'D1',raw:'LOW'}]}]}]});assert.equal(d.active.book.getRaw('D1'),'LOW');
});
test('v3 bounded repeat executes deterministically and previews safely',()=>{
 const d=make(),m={...base,steps:[{kind:'repeat',times:3,steps:[{kind:'replace',find:'A',replace:'AA',caseSensitive:true}]}]};
 const previous=JSON.stringify(d.toJSON());assert.ok(runActionMacro(d,m,{dryRun:true}).values>=1);assert.equal(JSON.stringify(d.toJSON()),previous);
 runActionMacro(d,m);assert.equal(d.active.book.getRaw('C1'),'AAAAAAAA');
});
test('v3 branch with out-of-range edit rolls back entire macro',()=>{
 const d=make(),before=JSON.stringify(d.toJSON());
 const m={...base,steps:[{kind:'sequence',range:'D1:D2',start:1,increment:1},{kind:'ifCell',ref:'B1',operator:'greaterThan',value:'0',thenSteps:[{kind:'cells',edits:[{ref:'ZZ999',raw:'bad'}]}]}]};
 assert.throws(()=>runActionMacro(d,m),/range|bound|exceed/i);assert.equal(JSON.stringify(d.toJSON()),before);
});
test('v3 invalid remote execution, unsafe script, infinite loops, excessive operations rejected',()=>{
 const bad=[{kind:'repeat',times:1000000,steps:[]},{kind:'javascript',code:'fetch("https://example.com")'},{kind:'sequence',range:'A1:A2',start:Infinity,increment:1},{kind:'ifCell',ref:'B1',operator:'eval',thenSteps:[]},{kind:'repeat',times:20,steps:[{kind:'clear',range:'A1:Z50'}]}];
 for(const step of bad)assert.throws(()=>validateMacro({...base,steps:[step]}));
});
test('v3 respects cross-sheet actions, and old v1/v2 recordings remain valid',()=>{
 const d=make();d.addSheet('Report');
 runActionMacro(d,{...base,steps:[{kind:'sequence',sheetName:'Report',range:'A1:A3',start:1,increment:1}]});
 assert.equal(d.sheets.find(s=>s.name==='Report').book.getRaw('A3'),'3');
 for(const version of [1,2]){const m=validateMacro({version,name:'old',sheetName:'Sheet 1',steps:[{kind:'cells',edits:[{ref:'H1',raw:'ok'}]}]});runActionMacro(d,m);}
 assert.equal(d.sheets.find(s=>s.name==='Sheet 1').book.getRaw('H1'),'ok');
});

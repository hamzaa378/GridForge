import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {validateMacro,runActionMacro} from '../dist/engine/macros.js';
const doc=()=>{const d=new SpreadsheetDocument();d.active.book.applyEdits([
 {ref:'A1',raw:'Name'},{ref:'B1',raw:'Quantity'},
 {ref:'A2',raw:'C'},{ref:'B2',raw:'3'},
 {ref:'A3',raw:'A'},{ref:'B3',raw:'1'},
 {ref:'A4',raw:'B'},{ref:'B4',raw:'2'},
 {ref:'D2',raw:'=B2*10'}]);
 d.addSheet('Summary');d.selectSheet(d.sheets[0].id);return d;};
test('v6 macros: cross-sheet edit + source-relative copying + sort + formatting + preview',()=>{
 const d=doc();const macro=validateMacro({version:2,name:'Quarterly report',sheetName:'Sheet 1',steps:[
 {kind:'sort',range:'A1:B4',byColumn:0,hasHeader:true},
 {kind:'fillDown',source:'D2',range:'D3:D4'},
 {kind:'styles',edits:[{ref:'A1',style:{bold:true,background:'#112233'}}]},
 {kind:'copy',sourceRange:'B2:B4',destination:'F2',valuesOnly:false},
 {kind:'cells',sheetName:'Summary',edits:[{ref:'A1',raw:'Revenue'},{ref:'B1',raw:'=SUM(1,2,3)'}]},
 {kind:'replace',sheetName:'Summary',find:'Revenue',replace:'Total'},
 ]});
 const before=JSON.stringify(d.toJSON()),pre=runActionMacro(d,macro,{dryRun:true});
 assert.ok(pre.values>=6);assert.equal(pre.dryRun,true);assert.deepEqual(pre.sheets,['Sheet 1','Summary']);
 assert.equal(JSON.stringify(d.toJSON()),before,'preview must not mutate original');
 const run=runActionMacro(d,macro);assert.equal(run.dryRun,false);
 const sheet=d.sheets.find(x=>x.name==='Sheet 1').book;
 assert.deepEqual(['A2','A3','A4'].map(x=>sheet.getRaw(x)),['A','B','C']);
 assert.deepEqual(['B2','B3','B4'].map(x=>sheet.getRaw(x)),['1','2','3']);
 assert.equal(sheet.getRaw('D3'),'=B3*10');assert.equal(sheet.getRaw('D4'),'=B4*10');
 assert.equal(sheet.getRaw('F2'),'1');assert.equal(sheet.getStyle('A1').bold,true);
 assert.equal(d.sheets.find(x=>x.name==='Summary').book.getRaw('A1'),'Total');
 assert.equal(d.sheets.find(x=>x.name==='Summary').book.getValue('B1').value,6);
 assert.equal(sheet.undo(),true);assert.equal(sheet.getRaw('A2'),'C');
});
test('v6 macros: invalid later step leaves both sheets unchanged',()=>{
 const d=doc(),before=JSON.stringify(d.toJSON());
 const macro={version:2,name:'Bad',sheetName:'Sheet 1',steps:[
 {kind:'cells',edits:[{ref:'A1',raw:'modified'}]},
 {kind:'cells',sheetName:'Summary',edits:[{ref:'B2',raw:'modified'}]},
 {kind:'clear',sheetName:'Summary',range:'ZZ999'},
 ]};
 assert.throws(()=>runActionMacro(d,macro),/boundary|exceeds/);
 assert.equal(JSON.stringify(d.toJSON()),before);
});
test('v6 macros: legacy v1 schema imports and preserves return structure',()=>{
 const d=doc(),macro=validateMacro({version:1,name:'old',sheetName:'Sheet 1',steps:[{kind:'cells',edits:[{ref:'A1',raw:'Legacy'}]}]});
 assert.deepEqual(runActionMacro(d,macro),{values:1,styles:0});
 assert.equal(d.active.book.getRaw('A1'),'Legacy');
});
test('v6 macros: unsafe unknown actions, scripts, invalid edits, and large operations rejected',()=>{
 for(const step of [
  {kind:'javascript',source:'alert(1)'},{kind:'vba',source:'Shell "whoami"'},
  {kind:'cells',edits:[{ref:'A1',raw:3}]},{kind:'sort',range:'A1:B2',byColumn:99},
  {kind:'clear',range:'A1:C1001'},{kind:'replace',find:'',replace:'b'},
 ])assert.throws(()=>validateMacro({version:2,name:'unsafe',sheetName:'Sheet 1',steps:[step]}));
});
test('v6 macros: clear contents and styles without formula evaluation',()=>{
 const d=doc();d.active.book.patchStyles(['B2'],{bold:true});
 runActionMacro(d,{version:2,name:'clear',sheetName:'Sheet 1',steps:[{kind:'clear',range:'B2:B3',clearFormatting:true}]});
 assert.equal(d.active.book.getRaw('B2'),'');assert.deepEqual(d.active.book.getStyle('B2'),{});
 assert.equal(d.active.book.getRaw('B4'),'2');
});

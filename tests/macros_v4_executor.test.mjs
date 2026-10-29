import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {validateMacro,runActionMacro} from '../dist/engine/macros.js';
const head={version:4,name:'Macro Studio',sheetName:'Sheet 1',parameters:[]};
const setup=()=>{const d=new SpreadsheetDocument();d.active.book.applyEdits([{ref:'A1',raw:'Name'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'  aLICE  '},{ref:'A3',raw:'bob'},{ref:'A4',raw:'bob'},{ref:'B2',raw:'15'},{ref:'B3',raw:'7'},{ref:'B4',raw:'99'},{ref:'C2',raw:'=B2*2'}]);return d;};
test('v4 named numeric and text parameters expand only in data fields',()=>{
  const d=setup(),m={...head,parameters:[{name:'rate',type:'number',required:true},{name:'label',type:'text',required:false,defaultValue:'Ready'}],steps:[{kind:'cells',edits:[{ref:'D2',raw:'=B2*(1+{{rate}})'},{ref:'E2',raw:'{{label}}'}]}]};
  const prior=JSON.stringify(d.toJSON());
  const preview=runActionMacro(d,m,{dryRun:true,parameters:{rate:'0.2'}});
  assert.equal(preview.values,2);assert.equal(JSON.stringify(d.toJSON()),prior);
  runActionMacro(d,m,{parameters:{rate:'0.2',label:'Approved'}});
  assert.equal(d.active.book.getValue('D2').value,18);assert.equal(d.active.book.getRaw('E2'),'Approved');
  assert.throws(()=>runActionMacro(d,m,{parameters:{rate:'not a number'}}),/numeric/);
  assert.throws(()=>runActionMacro(d,m,{parameters:{rate:'0.2',evil:'x'}}),/Unknown macro parameter/);
});
test('v4 data cleaning transforms text but preserves formulas by default',()=>{
  const d=setup();runActionMacro(d,{...head,steps:[{kind:'cleanText',range:'A2:A4',operation:'proper'},{kind:'cleanText',range:'C2:C2',operation:'upper'}]});
  assert.equal(d.active.book.getRaw('A2'),'  Alice  ');assert.equal(d.active.book.getRaw('A3'),'Bob');assert.equal(d.active.book.getRaw('C2'),'=B2*2');
});
test('v4 conditional formatting matches computed numeric values and is undoable',()=>{
  const d=setup();runActionMacro(d,{...head,steps:[{kind:'formatIf',range:'B2:C3',operator:'greaterThan',value:'20',style:{background:'#ffeeaa',bold:true}}]});
  assert.equal(d.active.book.getStyle('C2').background,'#ffeeaa');assert.equal(d.active.book.getStyle('B2').background,undefined);
  assert.equal(d.active.book.undo(),true);assert.equal(d.active.book.getStyle('C2').background,undefined);
});
test('v4 removeDuplicates compacts selected rows, retains header, preserves styles',()=>{
  const d=setup();d.active.book.applyStyles([{ref:'B4',style:{bold:true}}]);
  const m={...head,steps:[{kind:'removeDuplicates',range:'A1:B4',hasHeader:true,keyColumns:[0]}]};
  const before=JSON.stringify(d.toJSON());assert.equal(runActionMacro(d,m,{dryRun:true}).values>0,true);assert.equal(JSON.stringify(d.toJSON()),before);
  runActionMacro(d,m);
  assert.equal(d.active.book.getRaw('A4'),'');assert.equal(d.active.book.getRaw('B4'),'');assert.equal(d.active.book.getRaw('A3'),'bob');
  assert.equal(d.active.book.getStyle('B4').bold,undefined);
  assert.equal(d.active.book.undo(),true);assert.equal(d.active.book.getRaw('A4'),'bob');assert.equal(d.active.book.getStyle('B4').bold,true);
});
test('v4 failure in later step rolls back earlier changes',()=>{
  const d=setup(),before=JSON.stringify(d.toJSON());
  const m={...head,steps:[{kind:'cleanText',range:'A2:A4',operation:'upper'},{kind:'copy',sourceRange:'A1:B4',destination:'AZ399'}]};
  assert.throws(()=>runActionMacro(d,m),/bounds|exceed|outside/i);assert.equal(JSON.stringify(d.toJSON()),before);
});
test('v4 rejects arbitrary code, unknown actions, invalid keys and oversized mutations',()=>{
 const bad=[{kind:'script',command:'calc.exe'},{kind:'cleanText',range:'A1:A2',operation:'javascript'},{kind:'formatIf',range:'A1',operator:'equals',value:'a',style:{background:'red'}},{kind:'removeDuplicates',range:'A1:B4',keyColumns:[5]},{kind:'repeat',times:20,steps:[{kind:'cleanText',range:'A1:Z50',operation:'trim'}]}];
 for(const step of bad)assert.throws(()=>validateMacro({...head,steps:[step]}));
});
test('v1-v3 schemas remain loadable after adding v4',()=>{
 for(const version of [1,2,3])assert.equal(validateMacro({version,name:'legacy',sheetName:'Sheet 1',steps:[{kind:'cells',edits:[{ref:'A1',raw:'x'}]}]}).version,version);
});

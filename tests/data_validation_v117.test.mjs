import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {dataValidationAt,dataValidationError,dataValidationValues,freshTools,validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

test('v11.7 list validation locates covered cells and rejects values outside the dropdown',()=>{
 const rule={start:'B2',end:'B4',type:'list',values:['Open','In progress','Done'],allowBlank:true};
 const clean=validateTools({dataValidations:[rule]},20,8);
 assert.equal(dataValidationAt('B3',clean.dataValidations)?.type,'list');
 assert.equal(dataValidationAt('C3',clean.dataValidations),undefined);
 assert.equal(dataValidationError('Open',rule),null);
 assert.equal(dataValidationError('',rule),null);
 assert.match(dataValidationError('Closed',rule),/allowed list/);
 assert.equal(dataValidationError('=A1',rule),null);
});

test('v11.7 whole and decimal validation enforce bounds and persist with worksheet JSON',()=>{
 const whole={start:'A1',end:'A3',type:'whole',operator:'between',min:1,max:10,allowBlank:false};
 const decimal={start:'C1',end:'C5',type:'decimal',operator:'between',min:0,max:1,allowBlank:true};
 assert.equal(dataValidationError('7',whole),null);
 assert.match(dataValidationError('',whole),/Blank values are not allowed/);
 assert.match(dataValidationError('7.5',whole),/whole number/);
 assert.match(dataValidationError('11',whole),/from 1 to 10/);
 assert.equal(dataValidationError('0.25',decimal),null);
 assert.equal(dataValidationError('',decimal),null);
 const doc=new SpreadsheetDocument();doc.active.tools.dataValidations=[whole,decimal];
 const reopened=SpreadsheetDocument.fromJSON(doc.toJSON());
 assert.deepEqual(reopened.active.tools.dataValidations,[whole,decimal]);
 assert.equal(freshTools().dataValidations.length,0);
});

test('v11.7 validation rules export to native XLSX worksheet metadata',async()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.dataValidations=[
  {start:'B2',end:'B8',type:'list',values:['Open','In progress','Done'],allowBlank:true},
  {start:'C2',end:'C8',type:'whole',min:1,max:100,allowBlank:false},
  {start:'D2',end:'D8',type:'decimal',min:0,allowBlank:true}
 ];
 const files=await zipRead(exportXlsx(doc)),sheet=files.get('xl/worksheets/sheet1.xml');
 assert.match(sheet,/<dataValidations count="3">/);
 assert.match(sheet,/<dataValidation type="list"[^>]*sqref="B2:B8"><formula1>&quot;Open,In progress,Done&quot;<\/formula1>/);
 assert.match(sheet,/<dataValidation type="whole" operator="between"[^>]*sqref="C2:C8"><formula1>1<\/formula1><formula2>100<\/formula2>/);
 assert.match(sheet,/<dataValidation type="decimal" operator="greaterThanOrEqual"[^>]*sqref="D2:D8"><formula1>0<\/formula1>/);
});

test('v11.7 validation metadata rejects invalid ranges, list limits and reversed bounds',()=>{
 assert.throws(()=>validateTools({dataValidations:[{start:'B4',end:'B2',type:'list',values:['Yes'],allowBlank:true}]},20,8),/range or type/);
 assert.throws(()=>validateTools({dataValidations:[{start:'A1',end:'A1',type:'list',values:[],allowBlank:true}]},20,8),/List validation/);
 assert.throws(()=>validateTools({dataValidations:[{start:'A1',end:'A1',type:'whole',min:9,max:1,allowBlank:true}]},20,8),/criteria that match/);
});

test('v11.8 range-sourced dropdowns read live nonblank values and normalize the source range',()=>{
 const doc=new SpreadsheetDocument();doc.active.book.applyEdits([{ref:'A1',raw:'East'},{ref:'A2',raw:'West'},{ref:'A3',raw:''},{ref:'A4',raw:'East'}],false);
 const rule={start:'B1',end:'B5',type:'list',sourceRange:'$A$1:$A$4',allowBlank:true,errorStyle:'stop'};
 const clean=validateTools({dataValidations:[rule]},doc.active.book.rows,doc.active.book.cols).dataValidations[0];
 assert.equal(clean.sourceRange,'A1:A4');
 assert.deepEqual(dataValidationValues(clean,doc.active.book),['East','West']);
 assert.equal(dataValidationError('West',clean,dataValidationValues(clean,doc.active.book)),null);
 assert.match(dataValidationError('South',clean,dataValidationValues(clean,doc.active.book)),/East, West/);
});

test('v11.8 validation prompt and alert settings persist and export with range formulas',async()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.dataValidations=[{start:'B2',end:'B8',type:'list',sourceRange:'A2:A4',allowBlank:false,errorStyle:'warning',promptTitle:'Pick a region',promptMessage:'Choose a listed region.',errorTitle:'Unknown region',errorMessage:'Use the dropdown.'}];
 const reopened=SpreadsheetDocument.fromJSON(doc.toJSON());
 assert.deepEqual(reopened.active.tools.dataValidations,doc.active.tools.dataValidations);
 const files=await zipRead(exportXlsx(doc)),sheet=files.get('xl/worksheets/sheet1.xml');
 assert.match(sheet,/type="list" errorStyle="warning" allowBlank="0" showErrorMessage="1" showInputMessage="1" promptTitle="Pick a region" prompt="Choose a listed region\." errorTitle="Unknown region" error="Use the dropdown\." sqref="B2:B8"/);
 assert.match(sheet,/<formula1>\$A\$2:\$A\$4<\/formula1>/);
});

test('v11.8 validation messages enforce lengths and unsupported alert styles are rejected',()=>{
 assert.throws(()=>validateTools({dataValidations:[{start:'A1',end:'A1',type:'list',values:['A'],allowBlank:true,promptMessage:'x'.repeat(256)}]},20,8),/prompt or alert/);
 assert.throws(()=>validateTools({dataValidations:[{start:'A1',end:'A1',type:'list',values:['A'],allowBlank:true,errorStyle:'continue'}]},20,8),/prompt or alert/);
});

test('v11.8 numeric validation supports all Excel comparison operators',()=>{
 const rule=(operator,min,max)=>({start:'A1',end:'A1',type:'whole',operator,min,...(max===undefined?{}:{max}),allowBlank:true});
 assert.equal(dataValidationError('5',rule('between',1,10)),null);
 assert.match(dataValidationError('5',rule('notBetween',1,10)),/below 1 or above 10/);
 assert.match(dataValidationError('4',rule('equalTo',5)),/exactly 5/);
 assert.equal(dataValidationError('4',rule('notEqualTo',5)),null);
 assert.equal(dataValidationError('5',rule('greaterThan',4)),null);
 assert.equal(dataValidationError('5',rule('greaterThanOrEqual',5)),null);
 assert.equal(dataValidationError('4',rule('lessThan',5)),null);
 assert.equal(dataValidationError('5',rule('lessThanOrEqual',5)),null);
});

test('v11.8 numeric comparison operators export as native XLSX operators',async()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.dataValidations=[
  {start:'A1',end:'A2',type:'whole',operator:'equalTo',min:4,allowBlank:true},
  {start:'B1',end:'B2',type:'decimal',operator:'notEqualTo',min:1.5,allowBlank:true},
  {start:'C1',end:'C2',type:'whole',operator:'notBetween',min:2,max:8,allowBlank:true},
  {start:'D1',end:'D2',type:'decimal',operator:'lessThan',min:1,allowBlank:true}
 ];
 const files=await zipRead(exportXlsx(doc)),sheet=files.get('xl/worksheets/sheet1.xml');
 assert.match(sheet,/type="whole" operator="equal"[^>]*sqref="A1:A2"><formula1>4<\/formula1>/);
 assert.match(sheet,/type="decimal" operator="notEqual"[^>]*sqref="B1:B2"><formula1>1\.5<\/formula1>/);
 assert.match(sheet,/type="whole" operator="notBetween"[^>]*sqref="C1:C2"><formula1>2<\/formula1><formula2>8<\/formula2>/);
 assert.match(sheet,/type="decimal" operator="lessThan"[^>]*sqref="D1:D2"><formula1>1<\/formula1>/);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {dataValidationError,validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

const serial=iso=>Math.floor(Date.parse(`${iso}T00:00:00Z`)/86400000)+25569;

test('v11.11 date validation accepts ISO dates and applies Excel serial criteria',()=>{
 const rule={start:'A1',end:'A4',type:'date',operator:'between',min:serial('2025-01-01'),max:serial('2025-12-31'),allowBlank:true};
 assert.equal(dataValidationError('2025-04-20',rule),null);
 assert.match(dataValidationError('2026-01-01',rule),/from/);
 assert.match(dataValidationError('2025-02-30',rule),/valid date/);
 assert.throws(()=>validateTools({dataValidations:[{...rule,min:NaN}]},20,8),/criteria/i);
});

test('v11.11 text-length validation supports bounds and single-value operators',()=>{
 const rule={start:'B1',end:'B4',type:'textLength',operator:'between',min:2,max:5,allowBlank:true};
 assert.equal(dataValidationError('four',rule),null);
 assert.match(dataValidationError('toolong',rule),/character count/);
 assert.match(dataValidationError('x',rule),/character count/);
 assert.throws(()=>validateTools({dataValidations:[{...rule,min:1.5}]},20,8),/whole numbers/i);
});

test('v11.11 date and text-length rules persist in JSON and native XLSX validation records',async()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.dataValidations=validateTools({dataValidations:[
  {start:'A1',end:'A4',type:'date',operator:'greaterThanOrEqual',min:serial('2025-01-01'),allowBlank:true},
  {start:'B1',end:'B4',type:'textLength',operator:'between',min:2,max:24,allowBlank:false}
 ]},doc.active.book.rows,doc.active.book.cols).dataValidations;
 assert.deepEqual(SpreadsheetDocument.fromJSON(doc.toJSON()).active.tools.dataValidations,doc.active.tools.dataValidations);
 const parts=await zipRead(exportXlsx(doc)),sheet=parts.get('xl/worksheets/sheet1.xml');
 assert.match(sheet,/<dataValidation type="date" operator="greaterThanOrEqual"/);
 assert.match(sheet,/<dataValidation type="textLength" operator="between"/);
});

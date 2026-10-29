import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

test('v11.11 worksheet tabs can be reordered without changing the active sheet',()=>{
 const doc=new SpreadsheetDocument();const second=doc.addSheet('Second'),third=doc.addSheet('Third');
 doc.selectSheet(second.id);doc.reorderSheet(second.id,0);
 assert.deepEqual(doc.sheets.map(sheet=>sheet.name),['Second','Sheet 1','Third']);
 assert.equal(doc.activeId,second.id);
 assert.throws(()=>doc.reorderSheet(second.id,3),/position is out of range/i);
});

test('v11.11 worksheet tab colors validate, copy and persist in GridForge JSON',()=>{
 const doc=new SpreadsheetDocument();doc.setSheetTabColor(doc.activeId,'#4472c4');
 assert.throws(()=>doc.setSheetTabColor(doc.activeId,'blue'),/six-digit hex/i);
 const copy=doc.duplicateSheet();assert.equal(copy.tabColor,'#4472c4');
 const reopened=SpreadsheetDocument.fromJSON(doc.toJSON());
 assert.equal(reopened.sheets[0].tabColor,'#4472c4');
 assert.throws(()=>SpreadsheetDocument.fromJSON({...doc.toJSON(),sheets:doc.toJSON().sheets.map((sheet,index)=>index?sheet:{...sheet,tabColor:'bad'})}),/six-digit hex/i);
});

test('v11.11 worksheet tab colors export as native XLSX sheet properties',async()=>{
 const doc=new SpreadsheetDocument();doc.setSheetTabColor(doc.activeId,'#4472c4');
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/worksheets/sheet1.xml');
 assert.match(xml,/<sheetPr><tabColor rgb="FF4472C4"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
});

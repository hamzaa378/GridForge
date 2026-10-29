import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

test('v11.10 hide and unhide worksheets update tabs and persist in GridForge JSON',()=>{
 const doc=new SpreadsheetDocument();const second=doc.addSheet('Second');doc.hideSheet(second.id);
 assert.equal(doc.active.name,'Sheet 1');assert.equal(doc.sheets.find(sheet=>sheet.id===second.id).hidden,true);
 assert.throws(()=>doc.selectSheet(second.id),/Unhide this worksheet/);
 const reopened=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(reopened.sheets.find(sheet=>sheet.name==='Second').hidden,true);
 reopened.unhideSheet(reopened.sheets.find(sheet=>sheet.name==='Second').id);
 reopened.selectSheet(reopened.sheets.find(sheet=>sheet.name==='Second').id);
 assert.equal(reopened.active.name,'Second');
});

test('v11.10 a workbook always keeps one visible sheet and cannot delete its final visible tab',()=>{
 const doc=new SpreadsheetDocument();assert.throws(()=>doc.hideSheet(),/at least one worksheet visible/i);
 const hidden=doc.addSheet('Hidden');doc.hideSheet(hidden.id);
 assert.throws(()=>doc.hideSheet(),/at least one worksheet visible/i);
 assert.throws(()=>doc.deleteSheet(),/at least one worksheet visible/i);
});

test('v11.10 hidden worksheet visibility writes native XLSX workbook state',async()=>{
 const doc=new SpreadsheetDocument();const hidden=doc.addSheet('Hidden');doc.hideSheet(hidden.id);
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/workbook.xml');
 assert.match(xml,/<sheet name="Hidden" sheetId="2" state="hidden"/);
 assert.match(xml,/<workbookView activeTab="0"/);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

test('v11.11 hidden row and column state validates and persists in GridForge JSON',()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.hiddenRows=[2,4];doc.active.tools.hiddenColumns=[1,3];
 const reopened=SpreadsheetDocument.fromJSON(doc.toJSON());
 assert.deepEqual(reopened.active.tools.hiddenRows,[2,4]);
 assert.deepEqual(reopened.active.tools.hiddenColumns,[1,3]);
 assert.throws(()=>validateTools({hiddenRows:[0,1]},2,8),/at least one row must remain visible/i);
 assert.throws(()=>validateTools({hiddenColumns:[0,1]},20,2),/at least one column must remain visible/i);
});

test('v11.11 hidden rows and columns export using native XLSX flags',async()=>{
 const doc=new SpreadsheetDocument();doc.active.tools.hiddenRows=[1];doc.active.tools.hiddenColumns=[2];
 const parts=await zipRead(exportXlsx(doc)),sheet=parts.get('xl/worksheets/sheet1.xml');
 assert.match(sheet,/<row r="2" hidden="1"><\/row>/);
 assert.match(sheet,/<col min="3" max="3" width="17\.29" customWidth="1" hidden="1"\/>/);
});

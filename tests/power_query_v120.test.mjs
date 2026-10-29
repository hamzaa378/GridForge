import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {buildQueryRefresh,validateQuerySpec} from '../dist/engine/power_query.js';
import {validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

function sample(){const book=new Workbook(30,12);book.applyEdits([{ref:'A1',raw:' Name '},{ref:'B1',raw:'Amount'},{ref:'A2',raw:' alice '},{ref:'B2',raw:'12'},{ref:'A3',raw:'Bob'},{ref:'B3',raw:'7'}],false);return book;}
test('saved offline query replays cleanup and conversion steps into a separate load range',()=>{
 const book=sample(),spec=validateQuerySpec({id:'query-customer-cleanup',name:'Customer cleanup',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper','double']},book.rows,book.cols),result=buildQueryRefresh(book,spec);
 assert.equal(result.outputRange,'D1:E3');assert.deepEqual(['D1','E1','D2','E2','D3','E3'].map(ref=>book.getRaw(ref)),['','','','','','']);book.applyEdits(result.edits,false);assert.deepEqual(['D1','E1','D2','E2','D3','E3'].map(ref=>book.getValue(ref).value),['NAME','AMOUNT','ALICE',24,'BOB',14]);
});
test('query refresh picks up changed source values without changing the source during preview',()=>{
 const book=sample(),spec={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper'],lastOutputRange:'D1:E3'};book.applyEdits(buildQueryRefresh(book,spec).edits,false);book.applyEdits([{ref:'A2',raw:'  eve '}],false);const result=buildQueryRefresh(book,spec);assert.equal(book.getValue('A2').value,'  eve ');book.applyEdits(result.edits,false);assert.equal(book.getValue('D2').value,'EVE');
});
test('query definitions persist in GridForge workbook JSON and reject unsafe steps or overlap',()=>{
 const book=sample(),doc=new SpreadsheetDocument(book),query={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim']};doc.active.tools.queries=validateTools({queries:[query]},book.rows,book.cols).queries;const restored=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(restored.active.tools.queries[0]?.name,'Clean');assert.throws(()=>validateQuerySpec({...query,steps:['insert-sum']},book.rows,book.cols),/supported/);assert.throws(()=>validateQuerySpec({...query,destination:'B1'},book.rows,book.cols),/overlap/);
});
test('XLSX export refreshes saved query results without changing the open workbook',async()=>{
 const book=sample(),doc=new SpreadsheetDocument(book),spec={id:'query-clean',name:'Clean',sourceRange:'A1:B3',destination:'D1',steps:['trim','upper'],lastOutputRange:'D1:E3'};doc.active.tools.queries=validateTools({queries:[spec]},book.rows,book.cols).queries;book.applyEdits([{ref:'A2',raw:'  eve '}],false);
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/worksheets/sheet1.xml');assert.match(xml,/<c r="D2" t="inlineStr"><is><t xml:space="preserve">EVE<\/t><\/is><\/c>/);assert.equal(book.getValue('D2').value,null);assert.equal(book.getValue('A2').value,'  eve ');
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,display} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
const value=(book,ref)=>display(book.getValue(ref));
test('worksheet defined names resolve to cells/ranges and participate in recalculation dependencies',()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'10'},{ref:'A2',raw:'20'},{ref:'B1',raw:'=SUM(Sales)'},{ref:'B2',raw:'=TaxRate'}],false);
 book.setDefinedNames([{name:'Sales',ref:'A1:A2'},{name:'TaxRate',ref:'A1'}],false);
 assert.equal(value(book,'B1'),'30');assert.equal(value(book,'B2'),'10');
 assert.deepEqual(book.inspect('B1').dependencies,['A1','A2']);
 book.setDefinedNames([{name:'Sales',ref:'A1:A2'},{name:'TaxRate',ref:'A1'}]);assert.equal(value(book,'B2'),'10');
 book.setCell('A1','12');assert.equal(value(book,'B1'),'32');assert.equal(value(book,'B2'),'12');
});
test('defined names validate, persist as GridForge JSON and participate in undo/redo',()=>{
 const book=new Workbook(10,5);book.applyEdits([{ref:'A1',raw:'7'},{ref:'B1',raw:'=Amount*2'}],false);
 book.setDefinedNames([{name:'Amount',ref:'A1',comment:'Base amount'}]);assert.equal(value(book,'B1'),'14');
 const clone=Workbook.fromJSON(JSON.parse(JSON.stringify(book.toJSON())));assert.equal(value(clone,'B1'),'14');assert.deepEqual(clone.getDefinedNames(),[{name:'Amount',ref:'A1',comment:'Base amount'}]);
 assert.throws(()=>book.setDefinedNames([{name:'A1',ref:'A1'}]),/valid defined name/);
 assert.throws(()=>book.setDefinedNames([{name:'bad-name',ref:'A1'}]),/valid defined name/);
 assert.throws(()=>book.setDefinedNames([{name:'X',ref:'A1'},{name:'x',ref:'B1'}]),/already defined/);
 assert.throws(()=>book.setDefinedNames([{name:'FarAway',ref:'Z99'}]),/outside this worksheet/);
 assert.equal(book.undo(),true);assert.equal(book.getDefinedNames().length,0);assert.equal(value(book,'B1'),'#PARSE!');
 assert.equal(book.redo(),true);assert.equal(value(book,'B1'),'14');
});
test('worksheet-scoped names export through native XLSX definedNames',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.renameSheet(doc.activeId,'Sales Data');
 doc.active.book.applyEdits([{ref:'A1',raw:'5'},{ref:'A2',raw:'8'},{ref:'B1',raw:'=SUM(NetSales)'}],false);
 doc.active.book.setDefinedNames([{name:'NetSales',ref:'A1:A2',comment:'Exported total source'}],false);
 const archive=await zipRead(exportXlsx(doc)),xml=archive.get('xl/workbook.xml');assert.match(xml,/localSheetId="0"/);assert.match(xml,/NetSales/);assert.match(xml,/Sales Data/);
 assert.match(xml,/>'Sales Data'!\$A\$1:\$A\$2<\/definedName>/);
});

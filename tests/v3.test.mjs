import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {freshTools,validateTools,filterVisibleRows,ruleColor} from '../dist/engine/sheettools.js';
import {zipRead,zipWrite} from '../dist/engine/zip.js';
import {exportXlsx,xlsxFileName} from '../dist/engine/xlsx.js';
const make=()=>{
 const book=new Workbook();book.applyEdits([{ref:'A1',raw:'Item'},{ref:'A2',raw:'Widget'},{ref:'A3',raw:'Gadget'},{ref:'B2',raw:'20'},{ref:'B3',raw:'9'},{ref:'B4',raw:'=SUM(B2:B3)'}],false);
 book.patchStyles(['A1'],{bold:true,background:'#32659c'});
 const d=new SpreadsheetDocument(book);d.renameSheet(d.activeId,'Sales');return d;
};
test('ZIP writer and reader roundtrip UTF-8, verifies entry payloads',async()=>{
 const compressed=zipWrite({'xl/test.xml':'<tag>測試 — 💡</tag>','a.xml':'plain'});
 assert.equal(String.fromCharCode(...compressed.subarray(0,2)),'PK');
 const unpacked=await zipRead(compressed);
 assert.equal(unpacked.get('xl/test.xml'),'<tag>測試 — 💡</tag>');
 assert.equal(unpacked.get('a.xml'),'plain');
});
test('ZIP checksum catches corrupted bytes',async()=>{
 const bytes=zipWrite({'a.xml':'hello'});bytes[35]^=16;
 await assert.rejects(zipRead(bytes),/checksum/);
});
test('ZIP ignores arbitrary attachments but rejects traversal paths',async()=>{
 const output=await zipRead(zipWrite({'xl/test.xml':'safe','photo.png':'pretend-image'}));
 assert.equal(output.size,1);assert.equal(output.get('xl/test.xml'),'safe');
 assert.throws(()=>zipWrite({'../secret.xml':'bad'}),/Invalid ZIP/);
});
test('XLSX writer emits real multi-sheet Office XML with formulas and styles',async()=>{
 const doc=make();doc.active.tools.freezeRows=1;
 doc.active.tools.conditional.push({start:'B2',end:'B8',operator:'gt',threshold:12,color:'#ffe066'});
 doc.active.tools.filter={startRow:0,endRow:3,startCol:0,endCol:1,column:0,query:'Wid'};
 const sales=doc.active;
 doc.addSheet('Forecast');doc.active.book.setCell('A1','Forecast');
 const bytes=exportXlsx(doc);const parts=await zipRead(bytes);
 assert.equal(parts.size,7);assert.match(parts.get('xl/workbook.xml'),/name="Forecast"/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/<f>SUM\(B2:B3\)<\/f>/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/state="frozen"/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/<autoFilter/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/<conditionalFormatting/);
 assert.match(parts.get('xl/styles.xml'),/<dxfs count="1">/);
 assert.equal(sales.book.getValue('B4').value,29);
 assert.equal(xlsxFileName('Quarter / 4'),'Quarter _ 4.xlsx');
});
test('filter returns only matching values in data region while preserving source',()=>{
 const doc=make(),s=doc.active.book;
 const rows=filterVisibleRows(s,{startRow:0,endRow:3,startCol:0,endCol:1,column:0,query:'wid'});
 assert.deepEqual([...rows],[1]);assert.equal(s.getRaw('A3'),'Gadget');
});
test('conditional rules compare numeric values only and respect exact selected rectangle',()=>{
 const r=[{start:'B2',end:'C9',operator:'gt',threshold:15,color:'#ffcc00'}];
 assert.equal(ruleColor('B2',20,r),'#ffcc00');assert.equal(ruleColor('B2',15,r),null);
 assert.equal(ruleColor('A2',100,r),null);assert.equal(ruleColor('B2','20',r),null);
});
test('v2 JSON imports into v3 and serializes sheet feature state',()=>{
 const d=make();const old=d.toJSON();old.version=2;
 const decoded=SpreadsheetDocument.fromJSON(old);assert.equal(decoded.active.name,'Sales');
 decoded.active.tools.freezeCols=1;decoded.active.tools.conditional=[{start:'B2',end:'C6',operator:'lt',threshold:10,color:'#cc99ff'}];
 const restored=SpreadsheetDocument.fromJSON(JSON.parse(JSON.stringify(decoded.toJSON())));
 assert.equal(restored.active.tools.freezeCols,1);assert.equal(restored.active.tools.conditional.length,1);
 assert.equal(restored.active.book.getValue('B4').value,29);
});
test('malformed filter/rules are rejected on workbook restore',()=>{
 assert.throws(()=>validateTools({filter:{startRow:0,endRow:4000,startCol:0,endCol:1,column:0,query:''}},400,52),/Invalid/);
 assert.throws(()=>validateTools({conditional:[{start:'B2',end:'C6',operator:'xx',threshold:10,color:'#ffffff'}]},400,52),/Invalid/);
});

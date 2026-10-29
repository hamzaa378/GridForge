import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,display} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {resolveStructuredFormula} from '../dist/engine/tables.js';
import {expandTablesAfterEdit} from '../dist/engine/table_growth.js';
import {freshTools,validateTools,filterVisibleRows,ruleColor} from '../dist/engine/sheettools.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

const table={id:'table-orders',name:'Orders',range:'A1:C4',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false};
function book(){const b=new Workbook(50,12);b.applyEdits([
 {ref:'A1',raw:'Region'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Quantity'},
 ...[['North','100','2'],['South','200','3'],['North','300','4']].flatMap((row,i)=>row.map((raw,j)=>({ref:`${'ABC'[j]}${i+2}`,raw})))
 ]);return b;}
function chart(range=table.range){return {id:'chart-orders',kind:'column',title:'Revenue',range,x:30,y:50,width:400,height:260};}

test('v11 structured column reference in SUM preserves the editable raw formula',()=>{
 const b=book();b.setTableContext([table]);b.setCell('F1','=SUM(Orders[Revenue])');
 assert.equal(b.getRaw('F1'),'=SUM(Orders[Revenue])');assert.equal(b.getValue('F1').value,600);
 b.setCell('B3','400');assert.equal(b.getValue('F1').value,800);
});
test('v11 relative current-row table references calculate on each data row',()=>{
 const b=book();b.setTableContext([table]);b.setCell('D2','=[@Revenue]*[@Quantity]');
 // D2 sits outside the 3-column table, so current-row syntax is intentionally unavailable.
 assert.ok(b.getValue('D2').error);
 const wider={...table,range:'A1:D4'};b.setTableContext([wider]);
 assert.equal(b.getValue('D2').value,200);
});
test('v11 supported table qualifiers and quoted strings',()=>{
 const b=book();b.setTableContext([table]);
 assert.equal(resolveStructuredFormula('=SUM(Orders[[#All],[Revenue]])',b,[table],'F1'),'=SUM(B1:B4)');
 assert.equal(resolveStructuredFormula('=COUNT(Orders[[#Headers],[Revenue]])',b,[table],'F1'),'=COUNT(B1:B1)');
 assert.equal(resolveStructuredFormula('=COUNTA(Orders[#Data])',b,[table],'F1'),'=COUNTA(A2:C4)');
 assert.equal(resolveStructuredFormula('="Orders[Revenue]"&SUM(Orders[Revenue])',b,[table],'F1'),'="Orders[Revenue]"&SUM(B2:B4)');
});
test('v11 reject unknown structured column with explicit parse diagnostic',()=>{
 const b=book();b.setTableContext([table]);b.setCell('F1','=SUM(Orders[Missing])');
 assert.ok(b.getValue('F1').error);assert.match(b.getValue('F1').error.message,/Unknown table column/);
});
test('v11 auto expanding a table also expands its linked chart and updates structured formulas',()=>{
 const b=book(),tools=freshTools();tools.tables=[{...table}];tools.charts=[chart()];b.setTableContext(tools.tables);b.setCell('F1','=SUM(Orders[Revenue])');
 const edits=[{ref:'A5',raw:'West'},{ref:'B5',raw:'500'},{ref:'C5',raw:'5'}];b.applyEdits(edits);
 const changes=expandTablesAfterEdit(b,tools,edits);b.setTableContext(tools.tables);
 assert.deepEqual(changes.map(x=>[x.before,x.after,x.updatedCharts]),[['A1:C4','A1:C5',1]]);
 assert.equal(tools.charts[0].range,'A1:C5');assert.equal(b.getValue('F1').value,1100);
});
test('v11 table expansion respects contiguous row edits, overlaps and chart limits',()=>{
 const b=book(),tools=freshTools();tools.tables=[{...table},{...table,id:'table-next',name:'Another',range:'A5:C7'}];
 b.setCell('B5','500');assert.equal(expandTablesAfterEdit(b,tools,[{ref:'B5',raw:'500'}]).length,0);
 b.setCell('B9','900');assert.equal(expandTablesAfterEdit(b,tools,[{ref:'B9',raw:'900'}]).length,0);
 assert.equal(tools.tables[0].range,table.range);
});
test('v11 a legacy text filter still works',()=>{
 const b=book();const f={startRow:0,endRow:3,startCol:0,endCol:2,column:0,query:'nor'};
 assert.deepEqual([...filterVisibleRows(b,f)],[1,3]);
});
test('v11 exact multi-column filters combine using AND including blank and numeric display values',()=>{
 const b=book();const f={startRow:0,endRow:3,startCol:0,endCol:2,column:0,query:'',criteria:[{column:0,values:['North']},{column:2,values:['4']} ]};
 assert.deepEqual([...filterVisibleRows(b,f)],[3]);
 const t=validateTools({filter:f},b.rows,b.cols);assert.deepEqual(t.filter.criteria,f.criteria);
});
test('v11 filter validation rejects malformed and out-of-range criteria',()=>{
 const b=book();const f={startRow:0,endRow:3,startCol:0,endCol:2,column:0,query:'',criteria:[{column:5,values:['North']}]};
 assert.throws(()=>validateTools({filter:f},b.rows,b.cols),/Invalid filter values/);
 assert.throws(()=>validateTools({filter:{...f,criteria:[{column:0,values:['A']},{column:0,values:['B']}]}},b.rows,b.cols),/Invalid column-filter/);
});
test('v11 live 2-color scale interpolates from min to max and follows edited values',()=>{
 const b=book(),rule={start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#ffffff',color:'#000000'};
 assert.equal(ruleColor('B2',b.getValue('B2').value,[rule],b),'#ffffff');
 assert.equal(ruleColor('B3',b.getValue('B3').value,[rule],b),'#808080');
 assert.equal(ruleColor('B4',b.getValue('B4').value,[rule],b),'#000000');
 b.setCell('B2','300');assert.equal(ruleColor('B3',b.getValue('B3').value,[rule],b),'#ffffff');
});
test('v11 advanced features survive JSON save/open',()=>{
 const doc=new SpreadsheetDocument(book());doc.active.tools.tables=[{...table}];
 doc.active.book.setTableContext(doc.active.tools.tables);
 doc.active.book.setCell('F1','=SUM(Orders[Revenue])');
 doc.active.tools.filter={startRow:0,endRow:3,startCol:0,endCol:2,column:0,query:'',criteria:[{column:0,values:['North']}]};
 doc.active.tools.conditional=[{start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#fff2dc',color:'#1a9958'}];
 const loaded=SpreadsheetDocument.fromJSON(doc.toJSON());
 assert.equal(loaded.active.book.getValue('F1').value,600);
 assert.deepEqual([...filterVisibleRows(loaded.active.book,loaded.active.tools.filter)],[1,3]);
 assert.equal(loaded.active.tools.conditional[0].operator,'scale');
});
test('v11 XLSX writes real table formulas, multi-column autoFilters and colorScale XML',async()=>{
 const doc=new SpreadsheetDocument(book());doc.active.tools.tables=[{...table}];doc.active.book.setTableContext(doc.active.tools.tables);
 doc.active.book.setCell('F1','=SUM(Orders[Revenue])');
 doc.active.tools.filter={startRow:0,endRow:3,startCol:0,endCol:2,column:0,query:'',criteria:[{column:0,values:['North']},{column:2,values:['2','4']}]};
 doc.active.tools.conditional=[{start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#eef7ff',color:'#15784c'}];
 const binary=exportXlsx(doc),files=await zipRead(binary);
 const xml=files.get('xl/worksheets/sheet1.xml');
 assert.match(xml,/SUM\(Orders\[Revenue\]\)/);
 const nativeTable=files.get('xl/tables/table1.xml');
 assert.match(nativeTable,/colId="0"/);assert.match(nativeTable,/colId="2"/);assert.match(nativeTable,/<filter val="North"\/>/);
 assert.doesNotMatch(xml,/<autoFilter/);
 assert.match(xml,/type="colorScale"/);assert.match(xml,/<cfvo type="min"\/>/);
 // XLSX XML import requires the browser DOMParser; covered in Chromium UI smoke tests.
});

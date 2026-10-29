import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {validateTable} from '../dist/engine/tables.js';
import {propagateCalculatedColumns,fillCalculatedRows} from '../dist/engine/table_calculations.js';
import {reconcileCalculatedColumns} from '../dist/engine/table_calculations.js';

test('calculated column fills blanks, translates row formulas and preserves manual overrides',()=>{
 const book=new Workbook(20,10);book.applyEdits([
  {ref:'A1',raw:'Item'},{ref:'B1',raw:'Quantity'},{ref:'C1',raw:'Price'},{ref:'D1',raw:'Total'},
  {ref:'A2',raw:'Laptop'},{ref:'B2',raw:'2'},{ref:'C2',raw:'500'},{ref:'A3',raw:'Mouse'},{ref:'B3',raw:'4'},{ref:'C3',raw:'25'},{ref:'D3',raw:'101'},
  {ref:'A4',raw:'Keyboard'},{ref:'B4',raw:'3'},{ref:'C4',raw:'40'}
 ],false);
 const table=validateTable({id:'table-sales',name:'Sales',range:'A1:D4',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false},20,10);
 book.setTableContext([table]);
 const propagated=propagateCalculatedColumns(book,[table],[{ref:'D2',raw:'=[@Quantity]*[@Price]'}]);
 assert.deepEqual(propagated.map(x=>x.ref),['D2','D4']);
 book.applyEdits(propagated);book.setTableContext([table]);
 assert.equal(book.getValue('D2').value,1000);assert.equal(book.getValue('D3').value,101);assert.equal(book.getValue('D4').value,120);
 assert.deepEqual(table.calculatedColumns,[{column:3,formula:'=[@Quantity]*[@Price]',overrides:[1]}]);
});

test('new table rows receive formulas and overridden rows remain excluded',()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Qty'},{ref:'B1',raw:'Total'},{ref:'A2',raw:'2'},{ref:'B2',raw:'=A2*5'}],false);
 const table={id:'table-t',name:'T',range:'A1:B3',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false,calculatedColumns:[{column:1,formula:'=A2*5',overrides:[]}]};
 assert.deepEqual(fillCalculatedRows(book,table,2,2),[{ref:'B3',raw:'=A3*5'}]);
 table.calculatedColumns[0].overrides=[1];assert.deepEqual(fillCalculatedRows(book,table,2,2),[]);
});

test('propagated formulas are one undoable edit and metadata follows undo and redo',()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Qty'},{ref:'B1',raw:'Total'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'}],false);
 const table=validateTable({id:'table-u',name:'U',range:'A1:B3',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false},20,10);
 const edits=propagateCalculatedColumns(book,[table],[{ref:'B2',raw:'=A2*5'}]);book.applyEdits(edits);reconcileCalculatedColumns(book,[table]);
 assert.equal(book.getValue('B3').value,15);assert.equal(book.undo(),true);reconcileCalculatedColumns(book,[table]);assert.equal(book.getRaw('B2'),'');assert.equal(table.calculatedColumns.length,0);
 assert.equal(book.redo(),true);reconcileCalculatedColumns(book,[table]);assert.equal(book.getValue('B3').value,15);assert.equal(table.calculatedColumns.length,1);
});

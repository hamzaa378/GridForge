import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {buildPrintHtml} from '../dist/engine/print.js';
import {validateTools} from '../dist/engine/sheettools.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

test('page setup validates paper, scaling, margins and print areas',()=>{
 const tools=validateTools({pageSetup:{orientation:'landscape',paper:'a4',fitWidth:1,fitHeight:0,printArea:'$A$1:$B$3',printGridlines:true,printHeadings:false,marginTop:.5,marginBottom:.5,marginLeft:.4,marginRight:.4}},20,10);
 assert.deepEqual(tools.pageSetup,{orientation:'landscape',paper:'a4',fitWidth:1,fitHeight:0,printArea:'A1:B3',printGridlines:true,printHeadings:false,marginTop:.5,marginBottom:.5,marginLeft:.4,marginRight:.4});
 assert.throws(()=>validateTools({pageSetup:{...tools.pageSetup,paper:'folio'}},20,10),/page setup/i);
 assert.throws(()=>validateTools({pageSetup:{...tools.pageSetup,printArea:'A1:Z999'}},20,10),/print area/i);
});
test('print-ready view safely escapes cell text and applies sheet print settings',()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Name'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'<script>'},{ref:'B2',raw:'42'}],false);
 const html=buildPrintHtml(book,{orientation:'landscape',paper:'a4',fitWidth:1,fitHeight:0,printArea:'A1:B2',printGridlines:true,printHeadings:true,marginTop:.5,marginBottom:.5,marginLeft:.4,marginRight:.4},'Sales & Ops');
 assert.match(html,/size:A4 landscape/);assert.match(html,/Sales &amp; Ops/);assert.match(html,/&lt;script&gt;/);assert.match(html,/<th>1<\/th>/);assert.match(html,/border:1px solid/);
});
test('page setup and print area export as native worksheet and workbook XML',async()=>{
 const doc=new SpreadsheetDocument(),sheet=doc.active;sheet.book.applyEdits([{ref:'A1',raw:'Item'},{ref:'B1',raw:'Qty'},{ref:'A2',raw:'Pencil'},{ref:'B2',raw:'3'}],false);
 sheet.tools.pageSetup={orientation:'landscape',paper:'a4',fitWidth:1,fitHeight:0,printArea:'A1:B2',printGridlines:true,printHeadings:true,marginTop:.5,marginBottom:.6,marginLeft:.4,marginRight:.4};
 const parts=await zipRead(exportXlsx(doc)),ws=parts.get('xl/worksheets/sheet1.xml'),workbook=parts.get('xl/workbook.xml');
 assert.match(ws,/<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);assert.match(ws,/<printOptions gridLines="1" headings="1"\/>/);assert.match(ws,/<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"\/>/);assert.match(workbook,/_xlnm\.Print_Area/);assert.match(workbook,/'Sheet 1'!\$A\$1:\$B\$2/);
});

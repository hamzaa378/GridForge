import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,display} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {styledDisplay,cleanStyle} from '../dist/engine/format.js';
import {analyzeSelection,findCells,replaceCells,sortSelection} from '../dist/engine/operations.js';
const make=(values={})=>{const w=new Workbook(400,52);w.applyEdits(Object.entries(values).map(([ref,raw])=>({ref,raw:String(raw)})),false);return w;};
test('v2 cell styling is sparse and persists across JSON roundtrip',()=>{
 const w=make({A1:'12.45',B1:'=A1*2'});w.patchStyles(['A1','B1'],{bold:true,color:'#ff00AA',numberFormat:'currency'});
 const clone=Workbook.fromJSON(w.toJSON());assert.deepEqual(clone.getStyle('A1'),{bold:true,color:'#ff00AA',numberFormat:'currency'});
 assert.equal(styledDisplay(clone.getValue('B1'),clone.getStyle('B1')),'$24.90');assert.equal(clone.getRaw('B1'),'=A1*2');
});
test('formatting can be toggled off and number format reset',()=>{
 const w=make({A1:'3'});w.patchStyles(['A1'],{bold:true,italic:true,numberFormat:'percent',align:'right'});
 w.patchStyles(['A1'],{bold:false,numberFormat:'general',align:'left'});
 assert.deepEqual(w.getStyle('A1'),{italic:true});assert.equal(display(w.getValue('A1')),'3');
});
test('undo and redo include formatting transactions',()=>{
 const w=make({A1:'10'});w.patchStyles(['A1'],{bold:true});assert.equal(w.canUndo,true);
 w.undo();assert.deepEqual(w.getStyle('A1'),{});w.redo();assert.equal(w.getStyle('A1').bold,true);
});
test('copy/paste-style multi-edit transaction is one undoable action',()=>{
 const w=make({A1:'1',B1:'=A1+1'});w.applyTransaction([{ref:'A2',raw:'4'},{ref:'B2',raw:'=A2+1'}],[{ref:'A2',style:{color:'#ff00ff'}}]);
 assert.equal(display(w.getValue('B2')),'5');w.undo();assert.equal(w.getRaw('A2'),'');assert.equal(w.getRaw('B2'),'');assert.deepEqual(w.getStyle('A2'),{});
 w.redo();assert.equal(display(w.getValue('B2')),'5');assert.equal(w.getStyle('A2').color,'#ff00ff');
});
test('cell styles do not invalidate unrelated dependency caches',()=>{
 const w=make({A1:'4',B1:'=A1*2',C1:'=15'});w.setCell('A1','5');const before=w.lastStats;
 w.patchStyles(['B1'],{underline:true});assert.equal(w.lastStats,before);assert.equal(display(w.getValue('B1')),'10');
});
test('bad formatting is rejected on loading JSON',()=>{
 assert.throws(()=>cleanStyle({background:'url(javascript:evil)'}),/Invalid background/);
 assert.throws(()=>Workbook.fromJSON({rows:10,cols:3,cells:{A1:'5'},styles:{A1:{numberFormat:'nonsense'}}}),/Invalid numeric format/);
 assert.throws(()=>Workbook.fromJSON({rows:10,cols:3,cells:{},styles:{A1:{color:'red'}}}),/six-digit hex/);
});
test('number formats support percentage, currency, integer, and fixed decimals',()=>{
 const w=make({A1:'0.125',A2:'9999.997'});
 assert.equal(styledDisplay(w.getValue('A1'),{numberFormat:'percent'}),'12.5%');
 assert.equal(styledDisplay(w.getValue('A2'),{numberFormat:'number'}),'10,000.00');
 assert.equal(styledDisplay(w.getValue('A1'),{numberFormat:'currency'}),'$0.13');
 assert.equal(styledDisplay(w.getValue('A2'),{numberFormat:'integer'}),'10,000');
});
test('independent sheets have isolated cell values, calculation caches and formats',()=>{
 const doc=new SpreadsheetDocument(make({A1:'4',B1:'=A1*2'}));
 doc.renameSheet(doc.activeId,'Costs');const other=doc.addSheet('Revenue');other.book.setCell('A1','99');other.book.setCell('B1','=A1*2');
 assert.equal(display(doc.active.book.getValue('B1')),'198');
 assert.equal(display(doc.sheets[0].book.getValue('B1')),'8');
 assert.throws(()=>doc.renameSheet(doc.activeId,'COSTS'),/already exists/);
});
test('sheet duplication creates independent content including styles',()=>{
 const doc=new SpreadsheetDocument(make({A1:'25'}));doc.active.book.patchStyles(['A1'],{italic:true});
 const dup=doc.duplicateSheet();assert.equal(dup.name,'Sheet 1 copy');
 assert.equal(dup.book.getRaw('A1'),'25');assert.equal(dup.book.getStyle('A1').italic,true);
 dup.book.setCell('A1','100');assert.equal(doc.sheets[0].book.getRaw('A1'),'25');
});
test('workbook document export, import, and v1 legacy data migration',()=>{
 const doc=new SpreadsheetDocument(make({A1:'3'}));const second=doc.addSheet('Forecast');second.book.setCell('A1','10');
 const copy=SpreadsheetDocument.fromJSON(JSON.parse(JSON.stringify(doc.toJSON())));
 assert.equal(copy.sheets.length,2);assert.equal(copy.active.name,'Forecast');
 assert.equal(copy.sheets[0].book.getRaw('A1'),'3');assert.equal(copy.sheets[1].book.getRaw('A1'),'10');
 const old=SpreadsheetDocument.fromJSON({version:1,rows:10,cols:10,cells:{A1:'=2+3'}});
 assert.equal(old.active.book.getValue('A1').value,5);
});
test('deleting last worksheet and illegal name fails safely',()=>{
 const doc=new SpreadsheetDocument();assert.throws(()=>doc.deleteSheet(),/At least one sheet/);
 assert.throws(()=>doc.renameSheet(doc.activeId,'Bad/Name'),/Sheet name/);
 const second=doc.addSheet('Two');doc.deleteSheet(second.id);assert.equal(doc.sheets.length,1);
});
test('find and replace defaults to protecting formula sources',()=>{
 const w=make({A1:'Apple',A2:'pineapple',B1:'=UPPER(A1)',B2:'APPLE PIE'});
 assert.equal(findCells(w,'apple').length,3);
 const edits=replaceCells(w,'apple','pear');w.applyEdits(edits);
 assert.equal(w.getRaw('A1'),'pear');assert.equal(w.getRaw('A2'),'pinepear');assert.equal(w.getRaw('B2'),'pear PIE');
 assert.equal(w.getRaw('B1'),'=UPPER(A1)');assert.equal(display(w.getValue('B1')),'PEAR');
});
test('case sensitive search and explicit formula-source replacement',()=>{
 const w=make({A1:'Apple',A2:'APPLE',B1:'=UPPER(A1)'});
 assert.deepEqual(findCells(w,'Apple',true).map(x=>x.ref),['A1']);
 const edits=replaceCells(w,'A1','A2',true,true);w.applyEdits(edits);
 assert.equal(w.getRaw('B1'),'=UPPER(A2)');assert.equal(display(w.getValue('B1')),'APPLE');
});
test('sort selected rows in numeric order (not lexicographic) with stable rows',()=>{
 const w=make({A1:'10',B1:'ten',A2:'2',B2:'two',A3:'30',B3:'thirty'});
 w.applyEdits(sortSelection(w,{r1:0,r2:2,c1:0,c2:1},0));
 assert.deepEqual([w.getRaw('A1'),w.getRaw('A2'),w.getRaw('A3')],['2','10','30']);
 assert.deepEqual([w.getRaw('B1'),w.getRaw('B2'),w.getRaw('B3')],['two','ten','thirty']);
 w.applyEdits(sortSelection(w,{r1:0,r2:2,c1:0,c2:1},0,true));
 assert.deepEqual([w.getRaw('A1'),w.getRaw('A3')],['30','2']);
});
test('sorting rows repositions formulas and shifts their relative references',()=>{
 const w=make({A1:'30',B1:'=A1*2',A2:'10',B2:'=A2*2'});
 w.applyEdits(sortSelection(w,{r1:0,r2:1,c1:0,c2:1},0));
 assert.equal(w.getRaw('B1'),'=A1*2');assert.equal(w.getRaw('B2'),'=A2*2');
 assert.deepEqual([w.getValue('B1').value,w.getValue('B2').value],[20,60]);
});
test('selection analytics counts blanks errors and duplicate values',()=>{
 const w=make({A1:'2',B1:'2',C1:'text',A2:'5',B2:'=1/0'});
 const a=analyzeSelection(w,{r1:0,r2:1,c1:0,c2:2});
 assert.deepEqual({total:a.count,nums:a.numbers,sum:a.sum,blanks:a.blanks,errors:a.errors,duplicates:a.duplicates},{total:6,nums:3,sum:9,blanks:1,errors:1,duplicates:1});
 assert.equal(a.average,3);
});

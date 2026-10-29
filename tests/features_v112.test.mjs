import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook,display} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {runPivotJob,runPivotJobs} from '../dist/engine/pivot_job.js';
import {buildPivotSummary,pivotDateGroupLabel,pivotDateSerial,pivotSerialDateIso} from '../dist/engine/pivot.js';
import {chartSvg,histogramData,paretoData,extractChartData} from '../dist/engine/charts.js';
import {solveGoalSeek} from '../dist/engine/goal_seek.js';
import {validateTools} from '../dist/engine/sheettools.js';

test('cross-sheet references invalidate local dependents when the source changes',()=>{
  const doc=new SpreadsheetDocument(new Workbook(30,12));doc.renameSheet(doc.activeId,'Summary');
  const source=doc.addSheet('Data',new Workbook(30,12));source.book.applyEdits([{ref:'A1',raw:'4'}]);
  const target=doc.sheets[0].book;target.applyEdits([{ref:'A1',raw:'=Data!A1'},{ref:'B1',raw:'=A1+1'}]);
  assert.equal(display(target.getValue('B1')),'5');source.book.setCell('A1','11');assert.equal(display(target.getValue('B1')),'12');
});

test('document exposes cross-sheet recalculated cells so dependent PivotTables can refresh',()=>{
 const doc=new SpreadsheetDocument(new Workbook(30,12));doc.renameSheet(doc.activeId,'Summary');const summary=doc.active,source=doc.addSheet('Data',new Workbook(30,12));source.book.setCell('A1','4');summary.book.applyEdits([{ref:'A1',raw:'=Data!A1'},{ref:'B1',raw:'=A1+1'}]);assert.equal(summary.book.getValue('B1').value,5);
 source.book.setCell('A1','11');assert.equal(summary.book.getValue('B1').value,12);const changes=doc.consumeExternalInvalidations(),summaryChange=changes.find(change=>change.sheetName==='Summary');assert.ok(summaryChange);assert.ok(summaryChange.refs.includes('A1'));assert.ok(summaryChange.refs.includes('B1'));assert.deepEqual(doc.consumeExternalInvalidations(),[]);
});

test('quoted sheet names work and direct references update after rename and deletion',()=>{
  const doc=new SpreadsheetDocument(new Workbook(30,12));const first=doc.active;doc.renameSheet(first.id,'Main');
  const source=doc.addSheet('Quarter 1',new Workbook(30,12));source.book.setCell('C2','15');
  first.book.setCell('A1',"='Quarter 1'!C2");assert.equal(display(first.book.getValue('A1')),'15');
  doc.renameSheet(source.id,'Quarter 2');assert.equal(first.book.getRaw('A1'),"='Quarter 2'!C2");assert.equal(display(first.book.getValue('A1')),'15');
  doc.deleteSheet(source.id);assert.equal(first.book.getRaw('A1'),'=#REF!');assert.equal(display(first.book.getValue('A1')),'#REF!');
});

test('cross-sheet circular references report a cycle instead of recursing indefinitely',()=>{
  const doc=new SpreadsheetDocument(new Workbook(20,10));const one=doc.active;doc.renameSheet(one.id,'One');const two=doc.addSheet('Two',new Workbook(20,10));
  one.book.setCell('A1','=Two!A1');two.book.setCell('A1','=One!A1');assert.equal(one.book.getValue('A1').error?.code,'#CYCLE!');
});

test('dynamic array formulas spill, resize, clear old cells, and reject editing a spill child',()=>{
  const book=new Workbook(20,10);book.setCell('A1','=SEQUENCE(3,2)');
  assert.deepEqual(['A1','B1','A2','B2','A3','B3'].map(ref=>book.getValue(ref).value),[1,2,3,4,5,6]);assert.equal(book.getRaw('B1'),'');
  assert.throws(()=>book.setCell('B1','9'),/belongs to a spilled array/);
  book.setCell('A1','=SEQUENCE(1,2)');assert.equal(book.getValue('A2').value,null);assert.equal(book.getValue('B1').value,2);
});

test('blocked spill formulas return #SPILL! and recalculate when the blocking cell clears',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'D1',raw:'blocked'},{ref:'C1',raw:'=SEQUENCE(2,2)'}]);
  assert.equal(book.getValue('C1').error?.code,'#SPILL!');book.setCell('D1','');assert.equal(book.getValue('D1').value,2);assert.equal(book.getValue('C1').value,1);
});

test('spill reference operator returns the full local and cross-sheet arrays',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'=SEQUENCE(2,2)'},{ref:'D1',raw:'=SUM(A1#)'}]);assert.equal(book.getValue('D1').value,10);
  const doc=new SpreadsheetDocument(book),target=doc.addSheet('Remote',new Workbook(20,10));target.book.setCell('A1','=SEQUENCE(2,2)');book.setCell('D2','=SUM(Remote!A1#)');assert.equal(book.getValue('D2').value,10);
});

test('pivot summary groups row and column fields and emits grand totals',()=>{
  const book=new Workbook(30,12);book.applyEdits([
    {ref:'A1',raw:'Region'},{ref:'B1',raw:'Quarter'},{ref:'C1',raw:'Sales'},
    {ref:'A2',raw:'West'},{ref:'B2',raw:'Q1'},{ref:'C2',raw:'10'},
    {ref:'A3',raw:'West'},{ref:'B3',raw:'Q2'},{ref:'C3',raw:'12'},
    {ref:'A4',raw:'East'},{ref:'B4',raw:'Q1'},{ref:'C4',raw:'7'}
  ],false);
  const pivot=buildPivotSummary(book,{id:'pivot-sales',sourceRange:'A1:C4',destination:'E1',rowField:0,columnField:1,valueField:2,aggregation:'sum'});
  book.applyEdits(pivot.edits,false);
  assert.deepEqual(['E1','F1','G1','H1','E2','F2','G2','H2','E3','F3','G3','H3','E4','F4','G4','H4'].map(ref=>book.getValue(ref).value),
    ['Region','Q1 · Sum of Sales','Q2 · Sum of Sales','Grand Total','East',7,null,7,'West',10,12,22,'Grand Total',17,12,29]);
});
test('background PivotTable job rebuilds a summary from a serialized workbook snapshot',()=>{
 const doc=new SpreadsheetDocument(new Workbook(30,12)),book=doc.active.book;book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Sales'},
  {ref:'A2',raw:'West'},{ref:'B2',raw:'8'},
  {ref:'A3',raw:'East'},{ref:'B3',raw:'5'}
 ],false);
 const response=runPivotJob({requestId:'pivot-job-1',document:doc.toJSON(),destinationSheetId:doc.activeId,spec:{id:'pivot-job',sourceRange:'A1:B3',destination:'D1',rowField:0,valueField:1,aggregation:'sum'}});
 assert.equal(response.requestId,'pivot-job-1');assert.equal(response.error,undefined);assert.ok(response.result);
 assert.equal(response.result.rowCount,4);assert.deepEqual(response.result.edits.map(edit=>[edit.ref,edit.raw]),[['D1','Region'],['E1','Sum of Sales'],['D2','East'],['E2','5'],['D3','West'],['E3','8'],['D4','Grand Total'],['E4','13']]);
 const detail=runPivotJob({requestId:'pivot-job-details',document:doc.toJSON(),destinationSheetId:doc.activeId,detailRef:'E2',spec:{id:'pivot-job',sourceRange:'A1:B3',destination:'D1',rowField:0,valueField:1,aggregation:'sum'}});
 assert.deepEqual(detail.result?.detailSourceRows,[2]);
});
test('background PivotTable job reports a missing destination worksheet',()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));
 const response=runPivotJob({requestId:'pivot-job-missing',document:doc.toJSON(),destinationSheetId:'missing-sheet',spec:{id:'pivot-job-missing',sourceRange:'A1:B2',destination:'D1',rowField:0,valueField:1}});
 assert.equal(response.requestId,'pivot-job-missing');assert.match(response.error??'',/destination worksheet no longer exists/);
});
test('background PivotTable job resolves a source worksheet separately from its destination',()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.renameSheet(doc.activeId,'Report');const source=doc.addSheet('Data',new Workbook(20,10));source.book.applyEdits([
  {ref:'A1',raw:'Product'},{ref:'B1',raw:'Units'},
  {ref:'A2',raw:'Pen'},{ref:'B2',raw:'3'},
  {ref:'A3',raw:'Book'},{ref:'B3',raw:'4'}
 ],false);
 const response=runPivotJob({requestId:'pivot-job-cross-sheet',document:doc.toJSON(),destinationSheetId:doc.activeId,spec:{id:'pivot-job-cross-sheet',sourceSheet:'Data',sourceRange:'A1:B3',destination:'D1',rowField:0,valueField:1,aggregation:'sum'}});
 assert.equal(response.error,undefined);assert.deepEqual(response.result?.edits.map(edit=>[edit.ref,edit.raw]),[['D1','Product'],['E1','Sum of Units'],['D2','Book'],['E2','4'],['D3','Pen'],['E3','3'],['D4','Grand Total'],['E4','7']]);
});
test('PivotTable worker accepts source ranges larger than 12,000 cells up to the worksheet limit',()=>{
 const doc=new SpreadsheetDocument(new Workbook(1000,30));doc.renameSheet(doc.activeId,'Report');const reportId=doc.activeId,source=doc.addSheet('Data',new Workbook(1000,30)).book,edits=[];
 for(let column=0;column<13;column++)edits.push({ref:`${String.fromCharCode(65+column)}1`,raw:column===0?'Category':column===12?'Amount':`Field ${column+1}`});
 for(let row=2;row<=1000;row++)edits.push({ref:`A${row}`,raw:`Item ${Math.ceil((row-1)/2)}`},{ref:`M${row}`,raw:'1'});
 source.applyEdits(edits,false);
 const response=runPivotJob({requestId:'pivot-job-large-source',document:doc.toJSON(),destinationSheetId:reportId,spec:{id:'pivot-large-source',sourceSheet:'Data',sourceRange:'A1:M1000',destination:'O1',rowField:0,valueField:12,aggregation:'sum'}});
 assert.equal(response.error,undefined);assert.equal(response.result?.rowCount,502);assert.equal(response.result?.edits.length,1004);assert.ok(response.result?.edits.some(edit=>edit.ref==='P502'&&edit.raw==='999'));
});
test('PivotTable batch jobs share a snapshot and recalculate dependent summaries in order',()=>{
 const doc=new SpreadsheetDocument(new Workbook(40,12)),book=doc.active.book;book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Sales'},{ref:'A2',raw:'East'},{ref:'B2',raw:'5'},{ref:'A3',raw:'West'},{ref:'B3',raw:'8'},
  {ref:'C1',raw:'Linked sales'},{ref:'D1',raw:'Count'},{ref:'C2',raw:'=F2'},{ref:'D2',raw:'1'}
 ],false);
 const specs=[{id:'pivot-batch-first',sourceRange:'A1:B3',destination:'E1',rowField:0,valueField:1,aggregation:'sum'},{id:'pivot-batch-dependent',sourceRange:'C1:D2',destination:'G1',rowField:0,valueField:1,aggregation:'sum'}];
 const result=runPivotJobs({requestId:'pivot-batch-order',document:doc.toJSON(),destinationSheetId:doc.activeId,specs});
 assert.equal(result.error,undefined);assert.equal(result.results.length,2);assert.equal(result.results[1]?.result?.edits.find(edit=>edit.ref==='G2')?.raw,'5');assert.equal(result.results[1]?.result?.edits.find(edit=>edit.ref==='H2')?.raw,'1');
});
test('pivot summary applies a persisted exact-match report filter before grouping',()=>{
 const book=new Workbook(30,12);book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Quarter'},{ref:'C1',raw:'Sales'},
  {ref:'A2',raw:'West'},{ref:'B2',raw:'Q1'},{ref:'C2',raw:'10'},
  {ref:'A3',raw:'West'},{ref:'B3',raw:'Q2'},{ref:'C3',raw:'12'},
  {ref:'A4',raw:'East'},{ref:'B4',raw:'Q1'},{ref:'C4',raw:'7'}
 ],false);
 const spec={id:'pivot-filter',sourceRange:'A1:C4',destination:'E1',rowField:1,valueField:2,filterField:0,filterValue:'west',aggregation:'sum'},result=buildPivotSummary(book,spec);
 book.applyEdits(result.edits,false);assert.deepEqual(['E1','F1','E2','F2','E3','F3','E4','F4','E5','F5'].map(ref=>book.getValue(ref).value),['Region','west','Quarter','Sum of Sales','Q1',10,'Q2',12,'Grand Total',22]);
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push({...spec,lastOutputRange:result.outputRange});const restored=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(restored.active.tools.pivotTables[0]?.filterField,0);assert.equal(restored.active.tools.pivotTables[0]?.filterValue,'west');
 assert.throws(()=>buildPivotSummary(book,{...spec,filterField:1}));
});
test('Top value filters do not resurrect excluded items when show-empty-items is enabled',()=>{
 const book=new Workbook(30,12);book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Sales'},
  {ref:'A2',raw:'West'},{ref:'B2',raw:'20'},
  {ref:'A3',raw:'East'},{ref:'B3',raw:'10'},
  {ref:'A4',raw:'North'},{ref:'B4',raw:'5'}
 ],false);
 const result=buildPivotSummary(book,{id:'pivot-top-empty',sourceRange:'A1:B4',destination:'D1',rowField:0,valueField:1,aggregation:'sum',showEmptyItemsFields:[0],valueFilter:{field:0,measureField:1,direction:'top',count:1}});
 book.applyEdits(result.edits,false);
 assert.deepEqual(['D1','D2','D3'].map(ref=>book.getValue(ref).value),['Region','West','Grand Total']);
 assert.equal(book.getValue('E2').value,20);
});
test('Pivot slicer validation rejects manually grouped fields',()=>{
 const pivot={id:'pivot-manual-slicer',sourceRange:'A1:B4',destination:'D1',rowField:0,valueField:1,aggregation:'sum',manualGroups:[{field:0,groups:[{label:'Coastal',values:['West','East']}]}]},slicer={id:'slicer-manual-field',pivotId:pivot.id,field:0,caption:'Region',x:0,y:0,width:200,height:200};
 assert.throws(()=>validateTools({pivotTables:[pivot],pivotSlicers:[slicer]},30,12),/same ungrouped field/);
 const ungrouped={...pivot,id:'pivot-raw-slicer'};delete ungrouped.manualGroups;assert.equal(validateTools({pivotTables:[ungrouped],pivotSlicers:[{...slicer,id:'slicer-raw-field',pivotId:ungrouped.id}]},30,12).pivotSlicers.length,1);
});
test('Pivot date filters and grouping honor the Excel 1900 serial boundary',()=>{
 assert.equal(pivotDateSerial('1900-01-01'),1);
 assert.equal(pivotDateSerial('1900-02-28'),59);
 assert.equal(pivotDateSerial('1900-03-01'),61);
 assert.equal(pivotSerialDateIso(1),'1900-01-01');
 assert.equal(pivotSerialDateIso(60),'1900-02-29');
 assert.equal(pivotSerialDateIso(61),'1900-03-01');
 assert.equal(pivotDateGroupLabel(60,'days'),'1900-02-29');
 assert.equal(pivotDateGroupLabel(60,'months'),'Feb');
 assert.equal(pivotDateGroupLabel('2024-12-30','weeks'),'2025-W01');
 assert.equal(pivotDateGroupLabel('2025-01-05','weeks'),'2025-W01');
 assert.equal(pivotDateGroupLabel('2025-01-06','weeks'),'2025-W02');
 assert.equal(pivotDateGroupLabel('2021-01-01','weeks'),'2020-W53');
});

test('histogram data bins numeric inputs and renders a frequency chart',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'},{ref:'A4',raw:'4'},{ref:'A5',raw:'5'}],false);
  const data=histogramData(extractChartData(book,'A1:A5'));assert.equal(data.series[0].values.reduce((sum,n)=>sum+(n??0),0),5);assert.ok(data.rows>0&&data.rows<=20);
  assert.equal(histogramData(extractChartData(book,'A1:A5'),3).rows,3);assert.throws(()=>histogramData(extractChartData(book,'A1:A5'),51),/from 1 to 50/);
  const svg=chartSvg(book,{id:'chart-hist',kind:'histogram',title:'Distribution',range:'A1:A5',x:0,y:0,width:460,height:300});assert.match(svg,/<svg/);assert.match(svg,/<rect/);
});

test('waterfall charts render connected cumulative positive and negative contributions',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Item'},{ref:'B1',raw:'Change'},{ref:'A2',raw:'Start'},{ref:'B2',raw:'20'},{ref:'A3',raw:'Expense'},{ref:'B3',raw:'-7'},{ref:'A4',raw:'Growth'},{ref:'B4',raw:'12'}],false);
  const svg=chartSvg(book,{id:'chart-water',kind:'waterfall',title:'Bridge',range:'A1:B4',x:0,y:0,width:460,height:300,dataLabels:true});assert.match(svg,/#18865b/);assert.match(svg,/#d45d47/);assert.match(svg,/stroke-dasharray/);
});

test('Pareto chart orders categories by contribution and renders the cumulative percentage line',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Cause'},{ref:'B1',raw:'Count'},{ref:'A2',raw:'Late'},{ref:'B2',raw:'8'},{ref:'A3',raw:'Defect'},{ref:'B3',raw:'3'},{ref:'A4',raw:'Other'},{ref:'B4',raw:'1'}]);
  const data=paretoData(extractChartData(book,'A1:B4'));assert.deepEqual(data.labels,['Late','Defect','Other']);assert.deepEqual(data.series[1].values,[2/3,11/12,1]);
  const svg=chartSvg(book,{id:'chart-pareto',kind:'pareto',title:'Causes',range:'A1:B4',x:0,y:0,width:500,height:300,dataLabels:true});assert.match(svg,/<polyline/);assert.match(svg,/Cumulative %/);
  const invalid=new Workbook(20,10);invalid.applyEdits([{ref:'A1',raw:'Cause'},{ref:'B1',raw:'Count'},{ref:'A2',raw:'Bad'},{ref:'B2',raw:'-1'}]);assert.throws(()=>paretoData(extractChartData(invalid,'A1:B2')),/non-negative/);
});

test('Goal Seek previews a solution on a clone without changing the workbook',()=>{
  const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([{ref:'A1',raw:'4'},{ref:'B1',raw:'=A1*2'}]);
  const result=solveGoalSeek(doc,{sheetId:doc.activeId,targetCell:'B1',targetValue:20,changingCell:'A1'});assert.ok(Math.abs(result.changingValue-10)<1e-7);assert.ok(Math.abs(result.achievedValue-20)<1e-7);assert.equal(doc.active.book.getRaw('A1'),'4');
  assert.throws(()=>solveGoalSeek(doc,{sheetId:doc.activeId,targetCell:'A1',targetValue:20,changingCell:'B1'}),/target cell must contain a formula/);
});

test('box-and-whisker chart draws quartiles and flags Tukey outliers',()=>{
  const book=new Workbook(30,10);book.applyEdits([{ref:'A1',raw:'Sample A'},{ref:'B1',raw:'Sample B'},{ref:'A2',raw:'1'},{ref:'A3',raw:'2'},{ref:'A4',raw:'3'},{ref:'A5',raw:'4'},{ref:'A6',raw:'100'},{ref:'B2',raw:'10'},{ref:'B3',raw:'12'},{ref:'B4',raw:'14'},{ref:'B5',raw:'16'},{ref:'B6',raw:'18'}]);
  const svg=chartSvg(book,{id:'chart-box',kind:'box-whisker',title:'Samples',range:'A1:B6',x:0,y:0,width:500,height:300});assert.match(svg,/<rect/);assert.match(svg,/<circle/);assert.match(svg,/Sample A/);
});

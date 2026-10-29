import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {buildPivotSummary,pivotChartRange} from '../dist/engine/pivot.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {validateChart} from '../dist/engine/charts.js';
import {zipRead,zipWrite} from '../dist/engine/zip.js';
import {DOMParser} from '@xmldom/xmldom';
import {importXlsx} from '../dist/engine/xlsx.js';

globalThis.DOMParser=DOMParser;

function fixture(){
 const book=new Workbook(100,20);book.applyEdits(Object.entries({A1:'Region',B1:'Quarter',C1:'Sales',D1:'Channel',A2:'West',B2:'Q1',C2:'10',D2:'Web',A3:'West',B3:'Q2',C3:'12',D3:'Web',A4:'East',B4:'Q1',C4:'7',D4:'Retail',A5:'East',B5:'Q2',C5:'8',D5:'Web'}).map(([ref,raw])=>({ref,raw})),false);
 const spec={id:'pivot-native',sourceRange:'A1:D5',destination:'F1',rowField:0,columnField:1,valueField:2,filterField:3,filterValue:'Web',aggregation:'sum'},summary=buildPivotSummary(book,spec);book.applyEdits(summary.edits,false);spec.lastOutputRange=summary.outputRange;const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push(spec);return {doc,spec};
}
test('XLSX export writes native PivotTable, cache, source records, workbook and sheet relationships',async()=>{
 const {doc}=fixture(),parts=await zipRead(exportXlsx(doc));
 for(const path of ['xl/pivotCache/pivotCacheDefinition1.xml','xl/pivotCache/pivotCacheDefinition1.xml.rels','xl/pivotCache/pivotCacheRecords1.xml','xl/pivotTables/pivotTable1.xml','xl/pivotTables/_rels/pivotTable1.xml.rels'])assert.ok(parts.get(path),`missing ${path}`);
 const workbook=parts.get('xl/workbook.xml'),cache=parts.get('xl/pivotCache/pivotCacheDefinition1.xml'),records=parts.get('xl/pivotCache/pivotCacheRecords1.xml'),pivot=parts.get('xl/pivotTables/pivotTable1.xml');
 assert.match(workbook,/<pivotCaches><pivotCache cacheId="1" r:id="rId3"\/><\/pivotCaches>/);assert.match(cache,/worksheetSource ref="A1:D5" sheet="Sheet 1"/);assert.match(cache,/recordCount="4"/);assert.match(records,/count="4"/);assert.match(pivot,/axis="axisRow"/);assert.match(pivot,/axis="axisCol"/);assert.match(pivot,/axis="axisPage"/);assert.match(pivot,/subtotal="sum"/);assert.match(pivot,/pageField fld="3" item="0"/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/<pivotTableParts[^>]*count="1"/);assert.match(parts.get('xl/worksheets/_rels/sheet1.xml.rels'),/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/pivotTable"/);assert.match(parts.get('[Content_Types].xml'),/spreadsheetml\.pivotCacheDefinition\+xml/);assert.match(parts.get('[Content_Types].xml'),/spreadsheetml\.pivotTable\+xml/);
});

test('XLSX PivotTable export marks items hidden by caption filters',async()=>{
 const {doc,spec}=fixture();
 const filtered={...spec,id:'pivot-caption-filter',destination:'N1',labelFilters:[{field:0,operator:'equals',value:'west'}]};
 const summary=buildPivotSummary(doc.active.book,filtered);filtered.lastOutputRange=summary.outputRange;doc.active.book.applyEdits(summary.edits,false);doc.active.tools.pivotTables.push(filtered);
 const parts=await zipRead(exportXlsx(doc)),pivot=parts.get('xl/pivotTables/pivotTable2.xml');
 assert.ok(pivot);assert.match(pivot,/<pivotField axis="axisRow"[^>]*><items count="3"><item x="1" h="1"\/><item x="0"\/>/);assert.match(pivot,/<filter fld="0" type="captionEqual"[^>]*stringValue1="west"/);
});

test('PivotChart export emits its native PivotTable link metadata',async()=>{
 const {doc,spec}=fixture(),chartRange=pivotChartRange(spec),chart=validateChart({id:'chart-pivot-native',kind:'column',title:'Sales by region',range:chartRange,pivotId:spec.id,x:40,y:40,width:520,height:320,legend:'bottom',dataLabels:false},doc.active.book.rows,doc.active.book.cols);
 doc.active.tools.charts.push(chart);const bytes=await exportXlsx(doc),parts=await zipRead(bytes),chartXml=parts.get('xl/charts/chart1.xml'),pivotXml=parts.get('xl/pivotTables/pivotTable1.xml');
 assert.match(chartXml,/<c:pivotSource><c:name>GridForgePivot1<\/c:name><c:fmtId val="0"\/><\/c:pivotSource>/);assert.match(pivotXml,/<chartFormats count="1">/);
});
test('native slicer export writes linked cache, view, drawing anchor, and selected field items',async()=>{
 const {doc,spec}=fixture();spec.itemFilters=[{field:0,values:['West']}];
 doc.active.tools.pivotSlicers.push({id:'slicer-region',pivotId:spec.id,pivotIds:[spec.id],field:0,caption:'Region',x:480,y:40,width:220,height:280});
 const parts=await zipRead(exportXlsx(doc)),cache=parts.get('xl/slicerCaches/slicerCache1.xml'),view=parts.get('xl/slicers/slicer1.xml'),sheet=parts.get('xl/worksheets/sheet1.xml'),drawing=parts.get('xl/drawings/drawing1.xml'),workbook=parts.get('xl/workbook.xml'),rels=parts.get('xl/_rels/workbook.xml.rels');
 for(const [name,source] of Object.entries({'slicer cache':cache,'slicer view':view,'worksheet':sheet,'drawing':drawing,'workbook':workbook,'workbook relationships':rels}))assert.ok(source,`missing ${name}`);
 assert.match(cache,/sourceName="Region"/);assert.match(cache,/<pivotTable tabId="1" name="GridForgePivot1"\/>/);assert.match(cache,/<i x="0" s="1"\/><i x="1"\/>/);
 assert.match(view,/<slicer name="GridForgeSlicer1" cache="GridForgeSlicerCache1" caption="Region"/);assert.match(sheet,/<x14:slicer r:id="rId3"\/>/);assert.match(drawing,/<sle:slicer name="GridForgeSlicer1"\/>/);
 assert.match(workbook,/<x14:slicerCache r:id="rId4"\/>/);assert.match(rels,/relationships\/slicerCache" Target="slicerCaches\/slicerCache1.xml"/);assert.match(parts.get('xl/worksheets/_rels/sheet1.xml.rels'),/relationships\/slicer" Target="\.\.\/slicers\/slicer1.xml"/);
 assert.match(parts.get('[Content_Types].xml'),/application\/vnd\.ms-excel\.slicerCache\+xml/);assert.match(parts.get('[Content_Types].xml'),/application\/vnd\.ms-excel\.slicer\+xml/);
 for(const path of ['xl/slicerCaches/slicerCache1.xml','xl/slicers/slicer1.xml','xl/worksheets/sheet1.xml','xl/drawings/drawing1.xml','xl/workbook.xml']){
  const parsed=new DOMParser().parseFromString(parts.get(path),'application/xml');assert.equal(parsed.getElementsByTagName('parsererror').length,0,`${path} must be well-formed XML`);
 }
});
test('native slicer round-trip restores PivotTable selection, connections, and layout',async()=>{
 const {doc,spec}=fixture();spec.itemFilters=[{field:0,values:['West']}];const second={...spec,id:'pivot-connected',destination:'N1'};delete second.lastOutputRange;const secondSummary=buildPivotSummary(doc.active.book,second);doc.active.book.applyEdits(secondSummary.edits,false);second.lastOutputRange=secondSummary.outputRange;doc.active.tools.pivotTables.push(second);doc.active.tools.pivotSlicers.push({id:'slicer-region-roundtrip',pivotId:spec.id,pivotIds:[spec.id,second.id],field:0,caption:'Region',x:480,y:40,width:220,height:280},{id:'slicer-quarter-roundtrip',pivotId:spec.id,pivotIds:[spec.id,second.id],field:1,caption:'Quarter',x:720,y:40,width:220,height:280});
 const imported=await importXlsx(await exportXlsx(doc)),sheet=imported.document.active,slicer=sheet.tools.pivotSlicers[0],pivot=sheet.tools.pivotTables[0];
 assert.ok(slicer,JSON.stringify({warnings:imported.warnings}));assert.equal(sheet.tools.pivotSlicers.length,2);assert.equal(slicer.caption,'Region');assert.equal(slicer.field,0);assert.equal(slicer.x,480);assert.equal(slicer.y,40);assert.equal(slicer.pivotIds.length,2);assert.equal(sheet.tools.pivotSlicers[1].caption,'Quarter');assert.equal(sheet.tools.pivotSlicers[1].x,720);assert.ok(sheet.tools.pivotTables.every(item=>item.itemFilters?.find(filter=>filter.field===0)?.values[0]==='West'));
 assert.ok(!imported.warnings.some(warning=>warning.includes('native slicer was skipped')),JSON.stringify(imported.warnings));
});
test('native slicer round-trip preserves a selected blank item',async()=>{
 const {doc,spec}=fixture();doc.active.book.applyEdits([{ref:'A6',raw:''},{ref:'B6',raw:'Q3'},{ref:'C6',raw:'4'},{ref:'D6',raw:'Web'}],false);spec.sourceRange='A1:D6';spec.itemFilters=[{field:0,values:[null]}];const summary=buildPivotSummary(doc.active.book,spec);doc.active.book.applyEdits(summary.edits,false);spec.lastOutputRange=summary.outputRange;doc.active.tools.pivotSlicers.push({id:'slicer-blank-roundtrip',pivotId:spec.id,field:0,caption:'Region',x:400,y:40,width:220,height:280});
 const imported=await importXlsx(await exportXlsx(doc)),pivot=imported.document.active.tools.pivotTables[0];assert.deepEqual(pivot.itemFilters?.find(filter=>filter.field===0)?.values,[null]);
});
test('native slicer import reads Excel two-cell drawing anchors',async()=>{
 const {doc,spec}=fixture();doc.active.tools.pivotSlicers.push({id:'slicer-two-cell',pivotId:spec.id,field:0,caption:'Region',x:480,y:40,width:220,height:280});const parts=await zipRead(exportXlsx(doc)),drawing=parts.get('xl/drawings/drawing1.xml');
 parts.set('xl/drawings/drawing1.xml',drawing.replace('<xdr:absoluteAnchor><xdr:pos x="4572000" y="381000"/><xdr:ext cx="2095500" cy="2667000"/>','<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>') .replace('</xdr:absoluteAnchor>','</xdr:twoCellAnchor>'));
 const imported=await importXlsx(zipWrite(Object.fromEntries(parts))),slicer=imported.document.active.tools.pivotSlicers[0];assert.equal(slicer.x,448);assert.equal(slicer.y,26);assert.equal(slicer.width,448);assert.equal(slicer.height,234);
});
test('native PivotTable export imports back into an editable GridForge Pivot Summary',async()=>{
 const {doc}=fixture(),imported=await importXlsx(await exportXlsx(doc)),sheet=imported.document.active,pivot=sheet.tools.pivotTables[0];
 assert.ok(pivot,`the imported workbook should retain a PivotTable definition: ${JSON.stringify({warnings:imported.warnings,sheets:imported.document.sheets.map(item=>({name:item.name,pivots:item.tools.pivotTables}))})}`);
 assert.equal(pivot.sourceRange,'A1:D5');
 assert.deepEqual(pivot.rowFields??[pivot.rowField],[0]);
 assert.deepEqual(pivot.columnFields??[pivot.columnField],[1]);
 assert.deepEqual(pivot.filters?.[0]?.values,['Web']);
 assert.equal(sheet.book.getValue('I5').value,30);
 sheet.book.setCell('C2','100');const refreshed=buildPivotSummary(sheet.book,pivot);sheet.book.applyEdits(refreshed.edits,false);
 assert.equal(sheet.book.getValue('I5').value,120,'the imported PivotTable should recalculate from edited source data');
 assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('table-backed PivotTables export and import by table name and include the current table extent',async()=>{
 const {doc,spec}=fixture();doc.active.tools.tables.push({id:'table-pivot-source',name:'SalesData',range:'A1:D5',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false});spec.sourceTable='SalesData';
 doc.active.book.applyEdits([{ref:'A6',raw:'West'},{ref:'B6',raw:'Q3'},{ref:'C6',raw:'4'},{ref:'D6',raw:'Web'}],false);doc.active.tools.tables[0].range='A1:D6';
 const parts=await zipRead(exportXlsx(doc)),cache=parts.get('xl/pivotCache/pivotCacheDefinition1.xml');assert.match(cache,/<worksheetSource name="SalesData"\/>/);assert.doesNotMatch(cache,/<worksheetSource ref=/);
 const imported=await importXlsx(await exportXlsx(doc)),pivot=imported.document.active.tools.pivotTables[0];assert.equal(pivot.sourceTable,'SalesData');assert.equal(pivot.sourceRange,'A1:D6');assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('native PivotTables with a source on another worksheet round-trip as editable reports',async()=>{
 const {doc,spec}=fixture(),source=doc.active,destination=doc.addSheet('Report'),crossSheet={...spec,id:'pivot-cross-sheet',sourceSheet:source.name,destination:'A1'};delete crossSheet.lastOutputRange;const summary=buildPivotSummary(destination.book,crossSheet,source.book);destination.book.applyEdits(summary.edits,false);crossSheet.lastOutputRange=summary.outputRange;destination.tools.pivotTables.push(crossSheet);
 const imported=await importXlsx(await exportXlsx(doc)),report=imported.document.sheets.find(sheet=>sheet.name==='Report'),pivot=report?.tools.pivotTables[0];assert.ok(pivot,JSON.stringify({warnings:imported.warnings}));assert.equal(pivot.sourceSheet,'Sheet 1');assert.equal(pivot.sourceRange,'A1:D5');assert.equal(report.book.getValue('D5').value,30);assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('Top N and show-empty-items survive native PivotTable round-trip',async()=>{
 const book=new Workbook(100,20);book.applyEdits(Object.entries({A1:'Region',B1:'Sales',A2:'West',B2:'20',A3:'East',B3:'10',A4:'North',B4:'5'}).map(([ref,raw])=>({ref,raw})),false);
 const spec={id:'pivot-top-roundtrip',sourceRange:'A1:B4',destination:'D1',rowField:0,valueField:1,aggregation:'sum',showEmptyItemsFields:[0],valueFilter:{field:0,measureField:1,direction:'top',count:1}},result=buildPivotSummary(book,spec);book.applyEdits(result.edits,false);spec.lastOutputRange=result.outputRange;
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push(spec);const imported=await importXlsx(await exportXlsx(doc)),sheet=imported.document.active,pivot=sheet.tools.pivotTables[0];
 assert.ok(pivot,'the Top N PivotTable should import as an editable definition');
 assert.deepEqual(pivot.valueFilter,{field:0,measureField:1,measureIndex:0,direction:'top',count:1});
 assert.deepEqual(pivot.showEmptyItemsFields,[0]);
 assert.deepEqual(['D2','D3'].map(ref=>sheet.book.getValue(ref).value),['West','Grand Total']);
 assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('manual PivotTable groups survive native XLSX export and import',async()=>{
 const book=new Workbook(100,20);book.applyEdits(Object.entries({A1:'Region',B1:'Sales',A2:'West',B2:'20',A3:'East',B3:'10',A4:'North',B4:'5'}).map(([ref,raw])=>({ref,raw})),false);
 const spec={id:'pivot-manual-group-roundtrip',sourceRange:'A1:B4',destination:'D1',rowField:0,valueField:1,aggregation:'sum',manualGroups:[{field:0,groups:[{label:'Coastal',values:['West','East']}]}]},result=buildPivotSummary(book,spec);book.applyEdits(result.edits,false);spec.lastOutputRange=result.outputRange;
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push(spec);const imported=await importXlsx(await exportXlsx(doc)),sheet=imported.document.active,pivot=sheet.tools.pivotTables[0];
 assert.ok(pivot,'the manually grouped PivotTable should import as an editable definition');
 assert.deepEqual(pivot.manualGroups,spec.manualGroups);
 assert.deepEqual(['D2','D3'].map(ref=>sheet.book.getValue(ref).value),['North','Coastal']);
 assert.deepEqual(['E2','E3'].map(ref=>sheet.book.getValue(ref).value),[5,30]);
 assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('ISO week groups round-trip as native seven-day date groups',async()=>{
 const book=new Workbook(100,20);book.applyEdits(Object.entries({A1:'Date',B1:'Sales',A2:'2024-12-30',B2:'5',A3:'2025-01-05',B3:'7',A4:'2025-01-06',B4:'11'}).map(([ref,raw])=>({ref,raw})),false);
 const spec={id:'pivot-week-groups',sourceRange:'A1:B4',destination:'D1',rowField:0,valueField:1,aggregation:'sum',dateGroups:[{field:0,unit:'weeks'}]},summary=buildPivotSummary(book,spec);book.applyEdits(summary.edits,false);spec.lastOutputRange=summary.outputRange;
 assert.deepEqual(['D2','D3'].map(ref=>book.getValue(ref).value),['2025-W01','2025-W02']);assert.deepEqual(['E2','E3'].map(ref=>book.getValue(ref).value),[12,11]);
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push(spec);const parts=await zipRead(exportXlsx(doc)),cache=parts.get('xl/pivotCache/pivotCacheDefinition1.xml');assert.match(cache,/groupBy="days" groupInterval="7"/);
 const imported=await importXlsx(await exportXlsx(doc)),pivot=imported.document.active.tools.pivotTables[0];assert.deepEqual(pivot.dateGroups,[{field:0,unit:'weeks'}]);assert.deepEqual(['D2','D3'].map(ref=>imported.document.active.book.getValue(ref).value),['2025-W01','2025-W02']);assert.ok(!imported.warnings.some(warning=>warning.includes('PivotTable import skipped')),JSON.stringify(imported.warnings));
});
test('PivotTables with the same source and grouping reuse one native PivotCache',async()=>{
 const {doc,spec}=fixture(),second={...spec,id:'pivot-second',destination:'N1',lastOutputRange:undefined};delete second.lastOutputRange;const result=buildPivotSummary(doc.active.book,second);second.lastOutputRange=result.outputRange;doc.active.book.applyEdits(result.edits,false);doc.active.tools.pivotTables.push(second);
 const parts=await zipRead(exportXlsx(doc));assert.ok(parts.has('xl/pivotCache/pivotCacheDefinition1.xml'));assert.ok(parts.has('xl/pivotCache/pivotCacheRecords1.xml'));assert.ok(parts.has('xl/pivotTables/pivotTable2.xml'));assert.ok(!parts.has('xl/pivotCache/pivotCacheDefinition2.xml'));assert.equal((parts.get('xl/worksheets/sheet1.xml').match(/<pivotTablePart /g)??[]).length,2);assert.equal((parts.get('xl/workbook.xml').match(/<pivotCache /g)??[]).length,1);assert.equal((parts.get('xl/pivotTables/pivotTable2.xml').match(/cacheId="1"/)??[]).length,1);
});
test('compatible PivotTables on separate destination sheets share the source cache',async()=>{
 const {doc,spec}=fixture(),destination=doc.addSheet('Pivot 2'),second={...spec,id:'pivot-other-sheet',sourceSheet:'Sheet 1',destination:'A1'};delete second.lastOutputRange;const summary=buildPivotSummary(destination.book,second,doc.sheets[0].book);destination.book.applyEdits(summary.edits,false);second.lastOutputRange=summary.outputRange;destination.tools.pivotTables.push(second);
 const parts=await zipRead(exportXlsx(doc));assert.equal((parts.get('xl/workbook.xml').match(/<pivotCache /g)??[]).length,1);assert.match(parts.get('xl/pivotTables/pivotTable2.xml'),/cacheId="1"/);assert.match(parts.get('xl/pivotTables/_rels/pivotTable2.xml.rels'),/pivotCacheDefinition" Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/);assert.ok(!parts.has('xl/pivotCache/pivotCacheDefinition2.xml'));
});
test('PivotTables with different grouped-field order retain separate native caches',async()=>{
 const {doc,spec}=fixture();spec.manualGroups=[{field:0,groups:[{label:'West Region',values:['West']},{label:'East Region',values:['East']}]}];const second={...spec,id:'pivot-descending',destination:'N1',fieldSortOrders:{'0':'descending'}};delete second.lastOutputRange;const summary=buildPivotSummary(doc.active.book,second);doc.active.book.applyEdits(summary.edits,false);second.lastOutputRange=summary.outputRange;doc.active.tools.pivotTables.push(second);
 const parts=await zipRead(exportXlsx(doc));assert.ok(parts.has('xl/pivotCache/pivotCacheDefinition2.xml'));assert.equal((parts.get('xl/workbook.xml').match(/<pivotCache /g)??[]).length,2);assert.match(parts.get('xl/pivotTables/pivotTable2.xml'),/cacheId="2"/);
});
test('XLSX export refreshes pivot cells and cache from a private workbook snapshot',async()=>{
 const {doc}=fixture(),before=doc.active.book.getRaw('I5');doc.active.book.applyEdits([{ref:'C2',raw:'100'}],false);const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/worksheets/sheet1.xml');
 assert.match(xml,/<c r="I5"><v>120<\/v><\/c>/);assert.equal(doc.active.book.getRaw('I5'),before);assert.equal(doc.active.book.getRaw('C2'),'100');
});

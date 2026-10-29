import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
import {chartSvg,validateChart} from '../dist/engine/charts.js';
import {DOMParser} from '@xmldom/xmldom';
globalThis.DOMParser=DOMParser;
const w=new Workbook(200,20);
w.applyEdits(Object.entries({A1:'Quarter',B1:'Revenue',C1:'Expenses',A2:'Q1',B2:'100',C2:'30',A3:'Q2',B3:'=50*3',C3:'40',A4:'Q3',B4:'220',C4:'110'}).map(([ref,raw])=>({ref,raw})));
const base={id:'chart-sample',kind:'column',title:'Revenue & profits',range:'A1:C4',x:120,y:90,width:520,height:330,legend:'right',xAxisTitle:'Quarter',yAxisTitle:'USD',dataLabels:true};
test('Chart options are validated and saved through document JSON',()=>{
 assert.deepEqual(validateChart(base,200,20),base);
 assert.throws(()=>validateChart({...base,legend:'left'},200,20));
 assert.deepEqual(validateChart({...base,legend:'top'},200,20),{...base,legend:'top'});
 assert.throws(()=>validateChart({...base,xAxisTitle:'X'.repeat(71)},200,20));
 assert.throws(()=>validateChart({...base,dataLabels:'yes'},200,20));
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(base);const loaded=SpreadsheetDocument.fromJSON(doc.toJSON());assert.deepEqual(loaded.active.tools.charts[0],base);
});
test('Top chart legends render above the plot and export as native Excel chart metadata',async()=>{
 const chartSpec={...base,legend:'top'};const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(chartSpec);
 const svg=chartSvg(w,chartSpec);assert.match(svg,/y="56" width="10" height="10"/);assert.match(svg,/y="65"[^>]*>Revenue<\/text>/);
 const parts=await zipRead(exportXlsx(doc));assert.match(parts.get('xl/charts/chart1.xml'),/<c:legendPos val="t"\/>/);
});
for(const kind of ['column','bar','line','area','pie','scatter','doughnut','stacked-column','stacked-bar'])test(`Native DrawingML serialization for ${kind}`,async()=>{
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...base,kind});
 const parts=await zipRead(exportXlsx(doc));
 for(const name of ['xl/charts/chart1.xml','xl/drawings/drawing1.xml','xl/drawings/_rels/drawing1.xml.rels','xl/worksheets/_rels/sheet1.xml.rels'])assert.ok(parts.get(name),`missing ${name}`);
 const chart=parts.get('xl/charts/chart1.xml');
 assert.match(chart,/<c:chartSpace/);if(!['pie','doughnut'].includes(kind))assert.match(chart, /Quarter/);assert.match(chart, /Revenue &amp; profits/);assert.match(chart,/val="1"/);
 assert.ok(chart.includes('Q1')||kind==='scatter');
 assert.match(chart,/150/); // formula cached value is used, not the raw =50*3
 assert.ok(!chart.includes('50*3'));
 assert.match(chart,/<c:dLbls>/);
 if(kind==='doughnut')assert.match(chart,/<c:holeSize val="57"\/>/);
 if(kind.startsWith('stacked'))assert.match(chart,/<c:grouping val="stacked"\/>/);
 assert.match(parts.get('xl/worksheets/sheet1.xml'),/<drawing/);
 assert.match(parts.get('[Content_Types].xml'),/drawingml.chart\+xml/);
 assert.match(parts.get('xl/drawings/drawing1.xml'),/absoluteAnchor/);
 const svg=chartSvg(w,{...base,kind});assert.match(svg,/<svg/);
});
test('Histogram chart exports as a valid XLSX column chart fallback',async()=>{
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...base,kind:'histogram'});const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/charts/chart1.xml');assert.match(xml,/<c:barChart>/);assert.match(xml,/<c:barDir val="col"\/>/);
});
for(const [kind,layout] of [['waterfall','waterfall'],['pareto','paretoLine'],['box-whisker','boxWhisker'],['funnel','funnel'],['treemap','treemap'],['sunburst','sunburst']])test(`Specialized ${kind} chart exports as native ChartEx and imports back editable`,async()=>{
 const chart={...base,kind,range:'A1:B4'},doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(chart);const bytes=exportXlsx(doc),parts=await zipRead(bytes),xml=parts.get('xl/charts/chart1.xml');
 assert.match(xml,/<cx:chartSpace/);assert.match(xml,new RegExp(`layoutId="${layout}"`));assert.match(xml,/Revenue/);assert.match(parts.get('[Content_Types].xml'),/ms-office\.chartex\+xml/);assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'),/chartEx/);
 const imported=await importXlsx(bytes);assert.equal(imported.document.active.tools.charts[0]?.kind,kind);assert.equal(imported.document.active.tools.charts[0]?.range,'A1:B4');
});
test('Stock charts export as native editable OHLC chart metadata',async()=>{
 const stockBook=new Workbook(200,20);stockBook.applyEdits(Object.entries({A1:'Date',B1:'Open',C1:'High',D1:'Low',E1:'Close',A2:'2025-01-01',B2:'10',C2:'15',D2:'8',E2:'13',A3:'2025-01-02',B3:'13',C3:'16',D3:'11',E3:'12'}).map(([ref,raw])=>({ref,raw})));
 const doc=new SpreadsheetDocument(stockBook),spec={...base,kind:'stock',range:'A1:E3'};doc.active.tools.charts.push(spec);
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/charts/chart1.xml');
 assert.match(xml,/<c:stockChart>/);assert.equal((xml.match(/<c:ser>/g)??[]).length,4);assert.match(xml,/<c:hiLowLines\/>/);assert.match(xml,/<c:upDownBars>/);assert.match(xml,/\$B\$2:\$B\$3/);assert.match(xml,/\$E\$2:\$E\$3/);assert.ok(!xml.includes('<c:barChart>'));
 const noDate=new Workbook(200,20);noDate.applyEdits(Object.entries({A1:'Open',B1:'High',C1:'Low',D1:'Close',A2:'10',B2:'15',C2:'8',D2:'13',A3:'13',B3:'16',C3:'11',D3:'12'}).map(([ref,raw])=>({ref,raw})));const noDateDoc=new SpreadsheetDocument(noDate);noDateDoc.active.tools.charts.push({...spec,range:'A1:D3'});const noDateXml=(await zipRead(exportXlsx(noDateDoc))).get('xl/charts/chart1.xml');assert.match(noDateXml,/<c:cat><c:strLit>/);assert.ok(!noDateXml.includes("<c:cat><c:strRef><c:f>&apos;Sheet 1&apos;!$A$2:$A$3"));
 const hlc=new Workbook(200,20);hlc.applyEdits(Object.entries({A1:'Date',B1:'High',C1:'Low',D1:'Close',A2:'Day 1',B2:'15',C2:'8',D2:'13',A3:'Day 2',B3:'16',C3:'11',D3:'12'}).map(([ref,raw])=>({ref,raw})));const hlcDoc=new SpreadsheetDocument(hlc);hlcDoc.active.tools.charts.push({...spec,range:'A1:D3'});const hlcXml=(await zipRead(exportXlsx(hlcDoc))).get('xl/charts/chart1.xml');assert.match(hlcXml,/<c:stockChart>/);assert.equal((hlcXml.match(/<c:ser>/g)??[]).length,3);assert.match(hlcXml,/High/);assert.ok(!hlcXml.includes('Open'));assert.match(chartSvg(hlc,{...spec,range:'A1:D3'}),/<svg/);
});
test('Multiple worksheets and multiple charts have distinct OOXML paths and relationships',async()=>{
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(base,{...base,id:'chart-other',kind:'line',x:200});doc.addSheet('Costs',new Workbook(200,20));
 doc.active.book.applyEdits([{ref:'A1',raw:'T'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'Jan'},{ref:'B2',raw:'123'}]);doc.active.tools.charts.push({...base,id:'chart-third',range:'A1:B2'});
 const parts=await zipRead(exportXlsx(doc));
 assert.equal([...parts.keys()].filter(x=>/^xl\/charts\/chart\d+.xml$/.test(x)).length,3);
 assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'),/chart1\.xml/);
 assert.match(parts.get('xl/drawings/_rels/drawing1.xml.rels'),/chart2\.xml/);
 assert.match(parts.get('xl/drawings/_rels/drawing2.xml.rels'),/chart3\.xml/);
 assert.match(parts.get('xl/charts/chart3.xml'),/Costs/);
});
test('Charts preserve source data when formulas change, and exports get recalculated cached data',async()=>{
 const d1=new SpreadsheetDocument(w);d1.active.tools.charts.push(base);const before=(await zipRead(exportXlsx(d1))).get('xl/charts/chart1.xml');
 w.applyEdits([{ref:'B3',raw:'999'}]);const after=(await zipRead(exportXlsx(d1))).get('xl/charts/chart1.xml');
 assert.notEqual(before,after);assert.match(after,/<c:v>999<\/c:v>/);assert.match(after,/<c:f>&apos;Sheet 1&apos;!\$B\$2:\$B\$4<\/c:f>/);
});

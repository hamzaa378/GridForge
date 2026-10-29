import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {chartRange,chartSvg,extractChartData,validateChart,CHART_KINDS} from '../dist/engine/charts.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
const book=()=>{const w=new Workbook(100,16);w.applyEdits(Object.entries({A1:'Region',B1:'Sales',C1:'Costs',A2:'North',B2:'100',C2:'40',A3:'South',B3:'=50*3',C3:'80',A4:'West',B4:'-50',C4:'0'}).map(([ref,raw])=>({ref,raw})));return w;};
const spec=(kind='column')=>({id:'chart-revenue',kind,title:'Sales <region> & trends',range:'A1:C4',x:54,y:100,width:500,height:320});
test('Chart range validation accepts A1 and absolute references, rejects reverse and excessive bounds',()=>{
 assert.deepEqual(chartRange(' $A$1:$C$4 ',100,16),{r1:0,r2:3,c1:0,c2:2});
 for(const bad of ['A1','C10:A1','A1:Z1000','A1:A300','A1:Q3','A1:C0','A1:C5!'])assert.throws(()=>chartRange(bad,100,16));
});
test('Source-linked chart data includes headers, formula values, negatives and zeros',()=>{
 const d=extractChartData(book(),'A1:C4');assert.deepEqual(d.labels,['North','South','West']);
 assert.equal(d.series[0].name,'Sales');assert.deepEqual(d.series[0].values,[100,150,-50]);
 assert.deepEqual(d.series[1].values,[40,80,0]);
});
test('Switch Row/Column transposes chart series and preserves orientation through editable XLSX',async()=>{
 const w=book(),normal=extractChartData(w,'A1:C4'),switched=extractChartData(w,'A1:C4',true);
 assert.deepEqual(switched.labels,['Sales','Costs']);
 assert.deepEqual(switched.series.map(s=>s.name),['North','South','West']);
 assert.deepEqual(switched.series.map(s=>s.values),[[100,40],[150,80],[-50,0]]);
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...spec(),switchRowsColumns:true});
 const archive=await zipRead(exportXlsx(doc)),xml=archive.get('xl/charts/chart1.xml');
 assert.match(xml,/<c:f>&apos;Sheet 1&apos;!\$B\$1:\$C\$1<\/c:f>/);assert.match(xml,/<c:f>&apos;Sheet 1&apos;!\$B\$2:\$C\$2<\/c:f>/);
});
for(const kind of CHART_KINDS)test(`${kind} SVG chart renders offline without HTML injection`,()=>{
 const svg=chartSvg(book(),spec(kind));assert.match(svg,/^<svg /);assert.match(svg,/<\/svg>$/);
 assert.ok(svg.includes('&lt;region&gt;'));assert.ok(!svg.includes('<region>'));
 assert.ok(svg.length>400,`expected actual drawn chart, got ${svg.length} bytes`);
});
test('Zero and missing data produce SVG safely; pie negative values excluded',()=>{
 const w=book();w.applyEdits([{ref:'B2',raw:''},{ref:'B3',raw:''},{ref:'B4',raw:''}]);
 assert.match(chartSvg(w,spec()),/Select a range containing numeric values|<rect/);
 const pie=chartSvg(book(),spec('pie'));assert.match(pie,/path|circle/);
});
test('Chart metadata validation rejects malicious or corrupt data',()=>{
 assert.deepEqual(validateChart(spec(),100,16),spec());
 for(const changes of [{id:'<script>'},{kind:'javascript'},{width:1},{x:-1},{range:'Z1:ZZ100'},{title:'x'.repeat(101)}])assert.throws(()=>validateChart({...spec(),...changes},100,16));
});
test('Charts save and reload with workbooks, survive duplication and do not erase values',()=>{
 const doc=new SpreadsheetDocument(book());doc.active.tools.charts.push(spec());const json=doc.toJSON();const loaded=SpreadsheetDocument.fromJSON(json);
 assert.equal(loaded.active.tools.charts.length,1);assert.equal(loaded.active.book.getRaw('B3'),'=50*3');
 loaded.duplicateSheet();assert.deepEqual(loaded.active.tools.charts,[spec()]);
 loaded.active.tools.charts[0].title='New';assert.equal(loaded.sheets[0].tools.charts[0].title,spec().title);
});
test('Styles font, wrap, border, strike survive JSON, undo and XLSX serialization',async()=>{
 const w=book();const style={fontFamily:'Georgia',fontSize:18,border:'all',wrap:true,strike:true,bold:true,color:'#cc4400'};
 w.patchStyles(['B2'],style);assert.deepEqual(w.getStyle('B2'),style);w.undo();assert.deepEqual(w.getStyle('B2'),{});w.redo();
 const loaded=SpreadsheetDocument.fromJSON(new SpreadsheetDocument(w).toJSON());assert.deepEqual(loaded.active.book.getStyle('B2'),style);
 const xlsx=exportXlsx(loaded);const parts=await zipRead(xlsx);const xml=parts.get('xl/styles.xml');assert.match(xml,/<sz val="18"\/>/);assert.match(xml,/Georgia/);assert.match(xml,/<strike\/>/);assert.match(xml,/wrapText="1"/);assert.match(xml,/borders count="3"/);
});

test('Recalculation of linked cells updates chart geometry without editing chart metadata',()=>{
 const w=book(),chart=spec('line');const before=chartSvg(w,chart);
 w.applyEdits([{ref:'B3',raw:'999'}]);const after=chartSvg(w,chart);
 assert.notEqual(before,after);assert.equal(extractChartData(w,chart.range).series[0].values[1],999);
});
test('XY scatter uses numerical X column even when first row contains text headers',()=>{
 const w=new Workbook(20,4);w.applyEdits([{ref:'A1',raw:'X'},{ref:'B1',raw:'Y'},{ref:'A2',raw:'10'},{ref:'B2',raw:'20'},{ref:'A3',raw:'20'},{ref:'B3',raw:'100'}]);
 const d=extractChartData(w,'A1:B3');assert.equal(d.series.length,2);
 assert.deepEqual(d.series[0].values,[10,20]);assert.deepEqual(d.series[1].values,[20,100]);
 assert.match(chartSvg(w,{...spec('scatter'),range:'A1:B3'}),/<circle/);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {chartRange,chartSvg,extractChartData,validateChart,CHART_KINDS,applySeriesPointColorEdits} from '../dist/engine/charts.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
const book=()=>{const w=new Workbook(100,16);w.applyEdits(Object.entries({A1:'Region',B1:'Sales',C1:'Costs',A2:'North',B2:'100',C2:'40',A3:'South',B3:'=50*3',C3:'80',A4:'West',B4:'-50',C4:'0'}).map(([ref,raw])=>({ref,raw})));return w;};
const spec=(kind='column')=>({id:'chart-revenue',kind,title:'Sales <region> & trends',range:'A1:C4',x:54,y:100,width:500,height:320});
test('Chart range validation accepts A1 and absolute references, rejects reverse and excessive bounds',()=>{
 assert.deepEqual(chartRange(' $A$1:$C$4 ',100,16),{r1:0,r2:3,c1:0,c2:2});
 for(const bad of ['A1','C10:A1','A1:Z1000','A1:A300','A1:Q3','A1:C0','A1:C5!'])assert.throws(()=>chartRange(bad,100,16));
});
test('Chart source capacity expands to 10,000 plotted points while bounding series and memory',()=>{
 const large=new Workbook(20_000,40),wide={id:'chart-large-source',kind:'line',title:'Large data',range:'A1:B10001',x:0,y:0,width:500,height:300};
 assert.deepEqual(chartRange('A1:B10001',large.rows,large.cols),{r1:0,r2:10_000,c1:0,c2:1});
 assert.doesNotThrow(()=>validateChart(wide,large.rows,large.cols));
 assert.throws(()=>chartRange('A1:B10002',large.rows,large.cols),/10,000 plotted points/);
 assert.throws(()=>chartRange('A1:Y418',large.rows,large.cols),/10,000 plotted points/);
 const cells=[];for(let col=0;col<25;col++){const letter=String.fromCharCode(65+col);cells.push({ref:`${letter}1`,raw:col===0?'Category':`S${col}`});cells.push({ref:`${letter}2`,raw:col===0?'Row 1':'1'});}
 large.applyEdits(cells,false);assert.equal(extractChartData(large,'A1:Y2').series.length,24);
 large.setCell('A2','1');assert.throws(()=>extractChartData(large,'A1:Y2'),/up to 24 series/);
});
test('SVG previews draw all 24 supported series instead of silently truncating after eight',()=>{
 const wide=new Workbook(8,32),edits=[];
 for(let col=0;col<25;col++){
  const name=String.fromCharCode(65+col);
  edits.push({ref:`${name}1`,raw:col===0?'Category':`Series ${col}`});
  edits.push({ref:`${name}2`,raw:col===0?'North':String(col)});
  edits.push({ref:`${name}3`,raw:col===0?'South':String(col+1)});
  edits.push({ref:`${name}4`,raw:col===0?'East':String(col+2)});
 }
 wide.applyEdits(edits,false);
 const svg=chartSvg(wide,{...spec('line'),range:'A1:Y4',legend:'none'});
 assert.equal((svg.match(/<polyline /g)??[]).length,24);
 const legend=chartSvg(wide,{...spec('line'),range:'A1:Y4',legend:'bottom'});
 assert.equal((legend.match(/>Series \d+<\/text>/g)??[]).length,24);
 const radar=chartSvg(wide,{...spec('radar'),range:'A1:Y4',legend:'none'});
 assert.equal((radar.match(/<polygon /g)??[]).length,24);
 const radarLegend=chartSvg(wide,{...spec('radar'),range:'A1:Y4',legend:'bottom'});
 assert.equal((radarLegend.match(/>Series \d+<\/text>/g)??[]).length,24);
 const custom=chartSvg(wide,{...spec('radar'),range:'A1:Y4',legend:'none',seriesOptions:[{color:'#ff0000'}]});
 assert.match(custom,/stroke="#ff0000"/);
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
test('Pie and doughnut per-point colors use the full bounded 10,000-point chart capacity',()=>{
 const colors=Array.from({length:10_000},(_,index)=>index%2?'#2563eb':'#dc5b73');
 assert.equal(validateChart({...spec('pie'),dataPointColors:colors},100,16).dataPointColors.length,10_000);
 assert.throws(()=>validateChart({...spec('doughnut'),dataPointColors:Array.from({length:10_001},()=>'#2563eb')},100,16),/10,000 six-digit hexadecimal colors/);
});
test('Per-series point colors validate independently and reject malformed palettes',()=>{
 const first=['#123456','#abcdef','#774411'],second=['#e04040','#40a080','#4040e0'],chart={...spec('line'),seriesOptions:[{dataPointColors:first},{dataPointColors:second}]};
 const validated=validateChart(chart,100,16);assert.deepEqual(validated.seriesOptions?.map(options=>options.dataPointColors),[first,second]);
 assert.throws(()=>validateChart({...spec('line'),seriesOptions:[{dataPointColors:['red']}]},100,16),/six-digit hexadecimal colors/);
 assert.throws(()=>validateChart({...spec('line'),seriesOptions:[{dataPointColors:Array.from({length:10_001},()=>'#2563eb')}]},100,16),/10,000 six-digit hexadecimal colors/);
});
test('Chart editor point-color edits preserve per-series settings and validate indexes',()=>{
 const original=[{color:'#123456',dataPointColors:['#abcdef','#fedcba']},{color:'#e04040'}],edits=[{seriesIndex:0,pointIndex:3,color:'#102030'},{seriesIndex:1,pointIndex:1,color:'#405060'}],result=applySeriesPointColorEdits(original,5,edits,['#4472c4','#ed7d31']);
 assert.deepEqual(result[0].dataPointColors,['#abcdef','#fedcba','#123456','#102030','#123456']);assert.deepEqual(result[1].dataPointColors,['#e04040','#405060','#e04040','#e04040','#e04040']);assert.deepEqual(original[0].dataPointColors,['#abcdef','#fedcba']);
 assert.throws(()=>applySeriesPointColorEdits(original,5,[{seriesIndex:2,pointIndex:0,color:'#ffffff'}],['#4472c4','#ed7d31']),/Invalid chart point-color edit/);assert.throws(()=>applySeriesPointColorEdits(original,5,[{seriesIndex:0,pointIndex:5,color:'#ffffff'}],['#4472c4','#ed7d31']),/Invalid chart point-color edit/);assert.throws(()=>applySeriesPointColorEdits(original,5,[{seriesIndex:0,pointIndex:0,color:'red'}],['#4472c4','#ed7d31']),/Invalid chart point-color edit/);
});
test('Pie chart SVG legends use the same custom colors as their slices',()=>{
 const svg=chartSvg(book(),{...spec('pie'),dataPointColors:['#123456','#abcdef','#774411']});
 assert.ok((svg.match(/fill="#123456"/g)??[]).length>=2,'first-slice color appears on both its slice and legend key');
 assert.ok((svg.match(/fill="#abcdef"/g)??[]).length>=2,'second-slice color appears on both its slice and legend key');
});
test('Column charts render custom colors by category and validate their point-color range',()=>{
 const colors=['#123456','#abcdef','#774411'],svg=chartSvg(book(),{...spec('column'),dataPointColors:colors});
 for(const color of colors)assert.match(svg,new RegExp(`fill="${color}"`));
 assert.deepEqual(validateChart({...spec('bar'),dataPointColors:colors},100,16).dataPointColors,colors);
 assert.deepEqual(validateChart({...spec('line'),dataPointColors:colors},100,16).dataPointColors,colors);
 const markerSvg=chartSvg(book(),{...spec('line'),dataPointColors:colors});for(const color of colors)assert.match(markerSvg,new RegExp(`fill="${color}"`));
 assert.throws(()=>validateChart({...spec('waterfall'),dataPointColors:colors},100,16),/supported chart type/);
 assert.throws(()=>validateChart({...spec('area'),dataPointColors:colors},100,16),/supported chart type/);
});
test('Pie chart SVG preview keeps 32 positive slices instead of truncating after 16',()=>{
 const w=new Workbook(40,8),edits=[{ref:'A1',raw:'Category'},{ref:'B1',raw:'Value'}],colors=Array.from({length:32},(_,index)=>`#${(0x100000+index).toString(16)}`);
 for(let index=0;index<32;index++){edits.push({ref:`A${index+2}`,raw:`Category ${index+1}`},{ref:`B${index+2}`,raw:String(index+1)});}
 w.applyEdits(edits,false);const svg=chartSvg(w,{...spec('pie'),range:'A1:B33',legend:'none',dataPointColors:colors});
 assert.equal(colors.filter(color=>svg.includes(`fill="${color}"`)).length,32);
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
 const xlsx=exportXlsx(loaded);const parts=await zipRead(xlsx);const xml=parts.get('xl/styles.xml');assert.match(xml,/<sz val="18"\/>/);assert.match(xml,/Georgia/);assert.match(xml,/<strike\/>/);assert.match(xml,/wrapText="1"/);assert.match(xml,/borders count="2"/);
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
test('XY scatter draws line-marker paths and applies gap, span, and zero blank policies',()=>{
 const w=new Workbook(20,4);w.applyEdits([{ref:'A1',raw:'X'},{ref:'B1',raw:'Y'},{ref:'A2',raw:'1'},{ref:'B2',raw:'2'},{ref:'A3',raw:'2'},{ref:'B3',raw:''},{ref:'A4',raw:'3'},{ref:'B4',raw:'6'}],false);
 const chart={...spec('scatter'),range:'A1:B4'},pathPoints=(svg)=>{const points=svg.match(/<polyline[^>]*points="([^"]+)"/);return points?.[1].trim().split(/\s+/)??[];};
 assert.equal(pathPoints(chartSvg(w,{...chart,displayBlanksAs:'gap'})).length,0,'gap leaves two disconnected single-point runs');
 assert.equal(pathPoints(chartSvg(w,{...chart,displayBlanksAs:'span'})).length,2,'span connects the points around the missing Y value');
 assert.equal(pathPoints(chartSvg(w,{...chart,displayBlanksAs:'zero'})).length,3,'zero plots the missing value at zero');
 const markerless=chartSvg(w,{...chart,displayBlanksAs:'span',seriesOptions:[{marker:false}]});assert.equal((markerless.match(/<circle /g)??[]).length,0);assert.equal(pathPoints(markerless).length,2,'hiding markers keeps the scatter line visible');
});
test('Radar previews leave blank gaps, span missing categories, or plot blanks at zero',()=>{
 const w=new Workbook(20,4);w.applyEdits([{ref:'A1',raw:'Category'},{ref:'B1',raw:'Values'},{ref:'A2',raw:'A'},{ref:'B2',raw:'2'},{ref:'A3',raw:'B'},{ref:'B3',raw:''},{ref:'A4',raw:'C'},{ref:'B4',raw:'6'},{ref:'A5',raw:'D'},{ref:'B5',raw:'8'}],false);
 const chart={...spec('radar'),range:'A1:B5',legend:'none'},gap=chartSvg(w,{...chart,displayBlanksAs:'gap'}),span=chartSvg(w,{...chart,displayBlanksAs:'span'}),zero=chartSvg(w,{...chart,displayBlanksAs:'zero'});
 assert.equal((gap.match(/<polygon points=/g)??[]).length,0);assert.equal((gap.match(/<polyline points=/g)??[]).length,1,'the valid categories on either side of the circular boundary stay connected');
 assert.equal((span.match(/<polygon points=/g)??[]).length,1,'span connects valid points around the missing category');
 assert.equal((zero.match(/<polygon points=/g)??[]).length,1,'zero plots all categories, using zero for the blank');
});

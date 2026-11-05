import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';
import {zipRead,zipWrite} from '../dist/engine/zip.js';
import {colToName} from '../dist/engine/parser.js';
import {chartSvg,validateChart,extractScatterPairData} from '../dist/engine/charts.js';
import {exportOds,importOds} from '../dist/engine/ods.js';
import {DOMParser} from '@xmldom/xmldom';
globalThis.DOMParser=DOMParser;
const w=new Workbook(200,20);
w.applyEdits(Object.entries({A1:'Quarter',B1:'Revenue',C1:'Expenses',A2:'Q1',B2:'100',C2:'30',A3:'Q2',B3:'=50*3',C3:'40',A4:'Q3',B4:'220',C4:'110'}).map(([ref,raw])=>({ref,raw})));
const base={id:'chart-sample',kind:'column',title:'Revenue & profits',range:'A1:C4',x:120,y:90,width:520,height:330,legend:'right',xAxisTitle:'Quarter',yAxisTitle:'USD',dataLabels:true};
test('Chart options are validated and saved through document JSON',()=>{
 assert.deepEqual(validateChart(base,200,20),base);
 assert.equal(validateChart({...base,legend:'left'},200,20).legend,'left');
 assert.deepEqual(validateChart({...base,legend:'top'},200,20),{...base,legend:'top'});
 assert.throws(()=>validateChart({...base,xAxisTitle:'X'.repeat(71)},200,20));
 assert.throws(()=>validateChart({...base,dataLabels:'yes'},200,20));
 assert.throws(()=>validateChart({...base,xAxisMin:10,xAxisMax:10},200,20),/minimum must be less than its maximum/);
 assert.throws(()=>validateChart({...base,yAxisMajorUnit:0},200,20),/major unit must be a positive finite number/);
 assert.throws(()=>validateChart({...base,yAxisLogBase:10,yAxisMin:0},200,20),/bounds must be positive/);
 assert.throws(()=>validateChart({...base,xAxisMin:10,xAxisMax:10},200,20),/minimum must be less than its maximum/);
 assert.throws(()=>validateChart({...base,yAxisMajorUnit:0},200,20),/major unit must be a positive finite number/);
 assert.throws(()=>validateChart({...base,yAxisLogBase:10,yAxisMin:0},200,20),/bounds must be positive/);
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(base);const loaded=SpreadsheetDocument.fromJSON(doc.toJSON());assert.deepEqual(loaded.active.tools.charts[0],base);
});
test('Top chart legends render above the plot and export as native Excel chart metadata',async()=>{
 const chartSpec={...base,legend:'top'};const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(chartSpec);
 const svg=chartSvg(w,chartSpec);assert.match(svg,/y="56" width="10" height="10"/);assert.match(svg,/y="65"[^>]*>Revenue<\/text>/);
 const parts=await zipRead(exportXlsx(doc));assert.match(parts.get('xl/charts/chart1.xml'),/<c:legendPos val="t"\/>/);
});
test('Column and bar category gap width validates and round-trips through native XLSX charts',async()=>{
 const spec={...base,categoryGapWidth:35,seriesOverlap:60},doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(spec);
 assert.deepEqual(validateChart(spec,200,20),spec);assert.throws(()=>validateChart({...spec,categoryGapWidth:501},200,20));assert.throws(()=>validateChart({...spec,categoryGapWidth:1.5},200,20));assert.throws(()=>validateChart({...spec,seriesOverlap:101},200,20));assert.throws(()=>validateChart({...spec,seriesOverlap:-101},200,20));
 const svg=chartSvg(w,spec);assert.match(svg,/<rect x="\d+(?:\.\d+)?"/);
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.match(xml,/<c:gapWidth val="35"\/>/);assert.match(xml,/<c:overlap val="60"\/>/);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.equal(restored.categoryGapWidth,35);assert.equal(restored.seriesOverlap,60);
});
test('Line series dash styles render and round-trip through native XLSX charts',async()=>{
 const spec={...base,kind:'line',seriesOptions:[{lineStyle:'dashDot',lineWidth:3.25,markerSize:12},{lineStyle:'dot',lineWidth:.75,markerSize:8}]},doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(spec);
 assert.deepEqual(validateChart(spec,200,20),spec);assert.throws(()=>validateChart({...spec,seriesOptions:[{lineStyle:'wavy'}]},200,20));assert.throws(()=>validateChart({...spec,seriesOptions:[{markerSize:73}]},200,20));assert.throws(()=>validateChart({...spec,seriesOptions:[{lineWidth:3.1}]},200,20));
 const svg=chartSvg(w,spec);assert.match(svg,/stroke-dasharray="6 3 1 3"/);assert.match(svg,/stroke-dasharray="1 3"/);
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.match(xml,/prstDash val="dashDot"/);assert.match(xml,/prstDash val="dot"/);assert.match(xml,/c:size val="12"/);assert.match(xml,/c:size val="8"/);assert.match(xml,/a:ln w="41275"/);assert.match(xml,/a:ln w="9525"/);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.deepEqual(restored.seriesOptions.map(option=>option.lineStyle),['dashDot','dot']);assert.deepEqual(restored.seriesOptions.map(option=>option.markerSize),[12,8]);assert.deepEqual(restored.seriesOptions.map(option=>option.lineWidth),[3.25,.75]);
});
test('Combo charts round-trip per-series type, secondary-axis choice, marker, and trendline',async()=>{
 const spec={...base,kind:'combo',secondaryYAxisTitle:'Margin %',seriesOptions:[
  {kind:'column',axis:'primary',color:'#2563eb'},
  {kind:'line',axis:'secondary',color:'#dc5b73',marker:true,markerStyle:'diamond',trendline:{type:'linear',showEquation:true,showR2:true}},
 ]};
 const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(spec);const bytes=exportXlsx(doc),parts=await zipRead(bytes),xml=parts.get('xl/charts/chart1.xml');
 assert.match(xml,/<c:barChart>/);assert.match(xml,/<c:lineChart>/);assert.match(xml,/<c:axId val="404"\/>/);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.ok(restored);
 assert.equal(restored.kind,'combo');assert.equal(restored.secondaryYAxisTitle,'Margin %');
 assert.deepEqual(restored.seriesOptions,[{kind:'column',color:'#2563eb'},spec.seriesOptions[1]]);
});
test('Native XLSX charts preserve up to 24 linked series across larger category ranges',async()=>{
 const book=new Workbook(400,32),edits=[];
 for(let col=0;col<=24;col++){const letter=String.fromCharCode(65+col);edits.push({ref:`${letter}1`,raw:col===0?'Category':`Series ${col}`});}
 for(let row=2;row<=301;row++){edits.push({ref:`A${row}`,raw:`Category ${row-1}`});for(let col=1;col<=24;col++)edits.push({ref:`${String.fromCharCode(65+col)}${row}`,raw:String(row*col)});}
 book.applyEdits(edits,false);const doc=new SpreadsheetDocument(book),chart={...base,range:'A1:Y301'};doc.active.tools.charts.push(chart);
 const parts=await zipRead(exportXlsx(doc)),xml=parts.get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:ser>/g)??[]).length,24);
 const restored=(await importXlsx(exportXlsx(doc))).document.active.tools.charts[0];assert.ok(restored);assert.equal(restored.range,'A1:Y301');
});
test('XLSX pie chart preserves more than 250 per-slice colors on import and export',async()=>{
 const book=new Workbook(400,4),edits=[{ref:'A1',raw:'Category'},{ref:'B1',raw:'Value'}],colors=[];
 for(let row=2;row<=301;row++){edits.push({ref:`A${row}`,raw:`Slice ${row-1}`},{ref:`B${row}`,raw:String(row)});colors.push(row%2?'#2563eb':'#dc5b73');}
 book.applyEdits(edits,false);const doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({...base,kind:'pie',range:'A1:B301',dataPointColors:colors});
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:dPt>/g)??[]).length,300);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.equal(restored.dataPointColors.length,300);assert.deepEqual(restored.dataPointColors,colors);
});
test('XLSX column charts preserve category-specific colors on native import and export',async()=>{
 const colors=['#2563eb','#dc5b73','#18865b'];const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...base,kind:'column',dataPointColors:colors});
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:dPt>/g)??[]).length,colors.length*2);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);
});
test('XLSX combo charts color columns and line markers and import them natively',async()=>{
 const colors=['#2563eb','#dc5b73','#18865b'],doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...base,kind:'combo',dataPointColors:colors,seriesOptions:[{kind:'column'},{kind:'line'}]});
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:dPt>/g)??[]).length,colors.length*2);assert.equal((xml.match(/<c:dPt><c:idx val="\d+"\/><c:marker>/g)??[]).length,colors.length);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);
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
test('Histogram chart exports as native ChartEx from a single numeric series',async()=>{
 const book=new Workbook(200,20);book.applyEdits(Object.entries({A1:'Range',B1:'Value',A2:'x',B2:'10',A3:'y',B3:'15',A4:'z',B4:'20'}).map(([ref,raw])=>({ref,raw})));const doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({...base,kind:'histogram',range:'A1:B4'});const bytes=exportXlsx(doc),parts=await zipRead(bytes),xml=parts.get('xl/charts/chart1.xml');assert.match(xml,/<cx:chartSpace/);assert.match(xml,/layoutId="clusteredColumn"/);assert.match(xml,/<cx:binning/);const restored=await importXlsx(bytes);assert.equal(restored.document.active.tools.charts[0]?.kind,'histogram');assert.equal(restored.document.active.tools.charts[0]?.range,'B1:B4');
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
test('Line chart point colors export on native markers and round-trip without changing series stroke',async()=>{
 const colors=['#123456','#abcdef','#774411'],chart={...base,kind:'line',dataPointColors:colors,seriesOptions:[{color:'#224466',marker:true,markerStyle:'plus',markerSize:7}]},doc=new SpreadsheetDocument(w);doc.active.tools.charts.push(chart);
 const bytes=exportXlsx(doc),parts=await zipRead(bytes),xml=parts.get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:marker><c:symbol val="plus"\/><c:size val="7"\/><c:spPr>/g)??[]).length,3);for(const color of colors)assert.match(xml,new RegExp(`<c:marker>[\\s\\S]*?<a:srgbClr val="${color.slice(1).toUpperCase()}"/>`));assert.match(xml,/<c:spPr><a:solidFill><a:srgbClr val="224466"\/><\/a:solidFill>/);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);assert.equal(restored.seriesOptions?.[0]?.color,'#224466');
});
test('XLSX line and radar charts retain independent point colors per series; scatter retains its colored series',async()=>{
 const first=['#123456','#abcdef','#774411'],second=['#e04040','#40a080','#4040e0'];for(const kind of ['line','radar']){const doc=new SpreadsheetDocument(w);doc.active.tools.charts.push({...base,kind,range:'A1:C4',seriesOptions:[{dataPointColors:first},{dataPointColors:second}]});
 const restored=(await importXlsx(await exportXlsx(doc))).document.active.tools.charts[0];assert.deepEqual(restored.seriesOptions?.[0]?.dataPointColors,first,kind);assert.deepEqual(restored.seriesOptions?.[1]?.dataPointColors,second,kind);}
 const scatterBook=new Workbook(20,8);scatterBook.applyEdits([{ref:'A1',raw:'X'},{ref:'B1',raw:'Y1'},{ref:'C1',raw:'Y2'},{ref:'A2',raw:'1'},{ref:'B2',raw:'10'},{ref:'C2',raw:'5'},{ref:'A3',raw:'2'},{ref:'B3',raw:'20'},{ref:'C3',raw:'15'},{ref:'A4',raw:'3'},{ref:'B4',raw:'30'},{ref:'C4',raw:'25'}],false);const scatter=new SpreadsheetDocument(scatterBook);scatter.active.tools.charts.push({...base,kind:'scatter',range:'A1:C4',seriesOptions:[{dataPointColors:first},{dataPointColors:second}]});const restoredScatter=(await importXlsx(await exportXlsx(scatter))).document.active.tools.charts[0];assert.deepEqual(restoredScatter.seriesOptions?.[0]?.dataPointColors,first,'scatter series 1');assert.deepEqual(restoredScatter.seriesOptions?.[1]?.dataPointColors,second,'scatter series 2');
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
test('XY scatter pairs use independent X columns and preserve them through XLSX and ODS',async()=>{
 const book=new Workbook(20,10);book.applyEdits([
  {ref:'A1',raw:'X East'},{ref:'B1',raw:'East'},{ref:'C1',raw:'X West'},{ref:'D1',raw:'West'},
  {ref:'A2',raw:'1'},{ref:'B2',raw:'10'},{ref:'C2',raw:'100'},{ref:'D2',raw:'50'},
  {ref:'A3',raw:'2'},{ref:'B3',raw:'20'},{ref:'C3',raw:'200'},{ref:'D3',raw:'60'},
  {ref:'A4',raw:'3'},{ref:'B4',raw:'15'},{ref:'C4',raw:'300'},{ref:'D4',raw:'80'},
 ],false);
 const chart={...base,kind:'scatter',range:'A1:D4',scatterPairs:true,xAxisMin:0,xAxisMax:400,xAxisMajorUnit:100,yAxisMin:0,yAxisMax:100,yAxisMajorUnit:25,seriesOptions:[{color:'#2563eb',errorBars:{type:'fixed',value:2,axis:'x'}},{color:'#dc5b73',trendline:{type:'linear',showEquation:true,showR2:true}}]},doc=new SpreadsheetDocument(book);doc.active.tools.charts.push(chart);
 const data=extractScatterPairData(book,chart.range);assert.deepEqual(data.scatterXColumns,[0,2]);assert.deepEqual(data.scatterXValues,[[1,2,3],[100,200,300]]);assert.deepEqual(data.series.map(series=>series.name),['East','West']);
 assert.deepEqual(validateChart(chart,20,10),chart);assert.throws(()=>validateChart({...chart,kind:'line'},20,10));
 const svg=chartSvg(book,chart);assert.match(svg,/<title>West · 100: 50<\/title>/);assert.match(svg,/<title>East · 1: 10<\/title>/);assert.match(svg,/y = /);
 const xlsx=exportXlsx(doc),xlsxXml=(await zipRead(xlsx)).get('xl/charts/chart1.xml');assert.match(xlsxXml,/<c:f>&apos;Sheet 1&apos;!\$A\$2:\$A\$4<\/c:f>/);assert.match(xlsxXml,/<c:f>&apos;Sheet 1&apos;!\$C\$2:\$C\$4<\/c:f>/);assert.match(xlsxXml,/<c:f>&apos;Sheet 1&apos;!\$D\$2:\$D\$4<\/c:f>/);assert.match(xlsxXml,/<c:errDir val="x"\/>/);
 const xlsxChart=(await importXlsx(xlsx)).document.active.tools.charts[0];assert.equal(xlsxChart?.scatterPairs,true);assert.equal(xlsxChart?.range,'A1:D4');assert.equal(xlsxChart?.xAxisMin,0);assert.equal(xlsxChart?.xAxisMax,400);assert.equal(xlsxChart?.xAxisMajorUnit,100);assert.equal(xlsxChart?.yAxisMin,0);assert.equal(xlsxChart?.yAxisMax,100);assert.equal(xlsxChart?.yAxisMajorUnit,25);assert.equal(xlsxChart?.seriesOptions?.[0]?.errorBars?.axis,'x');assert.equal(xlsxChart?.seriesOptions?.[1]?.trendline?.type,'linear');
 const ods=exportOds(doc),odsParts=await zipRead(ods),odsChart=Array.from(odsParts.entries()).find(([path])=>path.includes('/content.xml')&&path.includes('Object'))?.[1]??'';assert.match(odsChart,/chart:domain table:cell-range-address="local-table\.\$A\$2:\$A\$4"/);assert.match(odsChart,/chart:domain table:cell-range-address="local-table\.\$C\$2:\$C\$4"/);assert.match(odsChart,/chart:values-cell-range-address="local-table\.\$D\$2:\$D\$4"/);assert.match(odsChart,/chart:error-indicator chart:dimension="x"/);
 const odsResult=await importOds(ods),odsSpec=odsResult.document.active.tools.charts[0];assert.equal(odsSpec?.kind,'scatter');assert.equal(odsSpec?.scatterPairs,true);assert.equal(odsSpec?.xAxisMin,0);assert.equal(odsSpec?.xAxisMax,400);assert.equal(odsSpec?.xAxisMajorUnit,100);assert.equal(odsSpec?.yAxisMin,0);assert.equal(odsSpec?.yAxisMax,100);assert.equal(odsSpec?.yAxisMajorUnit,25);assert.equal(odsSpec?.seriesOptions?.[0]?.errorBars?.axis,'x');assert.equal(odsSpec?.seriesOptions?.[1]?.trendline?.type,'linear');
});
test('Hidden scatter columns exclude whole X/Y pairs without shifting neighboring series or styles',async()=>{
 const book=new Workbook(12,8);book.applyEdits([
  {ref:'A1',raw:'Label'},{ref:'B1',raw:'X East'},{ref:'C1',raw:'East'},{ref:'D1',raw:'X West'},{ref:'E1',raw:'West'},
  {ref:'A2',raw:'P1'},{ref:'B2',raw:'1'},{ref:'C2',raw:'10'},{ref:'D2',raw:'100'},{ref:'E2',raw:'50'},
  {ref:'A3',raw:'P2'},{ref:'B3',raw:'2'},{ref:'C3',raw:'20'},{ref:'D3',raw:'200'},{ref:'E3',raw:'60'},
 ],false);
 const doc=new SpreadsheetDocument(book);doc.active.tools.hiddenColumns=[2];doc.active.tools.hiddenRows=[1];
 const visible=extractScatterPairData(book,'A1:E3',false);
 assert.deepEqual(visible.scatterXColumns,[3]);assert.deepEqual(visible.scatterPairIndexes,[1]);assert.deepEqual(visible.scatterXValues,[[200]]);
 assert.deepEqual(visible.series.map(item=>item.name),['West']);assert.deepEqual(visible.series[0]?.values,[60]);assert.deepEqual(visible.labels,['P2']);
 doc.active.tools.hiddenRows=[];
 const chart={...base,kind:'scatter',range:'A1:E3',scatterPairs:true,includeHiddenData:false,seriesOptions:[{color:'#ff0000',markerSize:2},{color:'#00aa22',markerSize:10,trendline:{type:'linear',showEquation:true,showR2:true},errorBars:{type:'fixed',value:2,axis:'x'}}]},svg=chartSvg(book,chart);assert.match(svg,/West · 100: 50/);assert.doesNotMatch(svg,/East ·/);assert.match(svg,/stroke="#00aa22"/);assert.match(svg,/fill="#00aa22"/);assert.doesNotMatch(svg,/#[fF][fF]0000/);assert.match(svg,/y = /);
 doc.active.tools.charts.push(chart);const xlsxParts=await zipRead(exportXlsx(doc)),xlsxChart=xlsxParts.get('xl/charts/chart1.xml')??'';assert.equal((xlsxChart.match(/<c:ser>/g)??[]).length,2);assert.match(xlsxChart,/<c:f>&apos;Sheet 1&apos;!\$B\$2:\$B\$3<\/c:f>/);assert.match(xlsxChart,/<c:f>&apos;Sheet 1&apos;!\$C\$2:\$C\$3<\/c:f>/);assert.match(xlsxChart,/<c:f>&apos;Sheet 1&apos;!\$D\$2:\$D\$3<\/c:f>/);assert.match(xlsxChart,/<c:f>&apos;Sheet 1&apos;!\$E\$2:\$E\$3<\/c:f>/);assert.match(xlsxChart,/val="FF0000"/);assert.match(xlsxChart,/val="00AA22"/);assert.match(xlsxChart,/<c:trendline>/);assert.match(xlsxChart,/<c:errDir val="x"\/>/);assert.match(xlsxChart,/<c:plotVisOnly val="1"\/>/);const xlsxRestored=(await importXlsx(exportXlsx(doc))).document.active.tools.charts[0];assert.equal(xlsxRestored?.scatterPairs,true);assert.notEqual(xlsxRestored?.includeHiddenData,true);assert.equal(xlsxRestored?.seriesOptions?.[1]?.color,"#00aa22");const odsParts=await zipRead(exportOds(doc)),chartPath=[...odsParts.keys()].find(path=>path.startsWith('Object ')&&path.endsWith('/content.xml')),chartXml=chartPath?odsParts.get(chartPath)??'':'';
 assert.match(chartXml,/svg:stroke-color="#00aa22"/);assert.match(chartXml,/chart:error-indicator[^>]*chart:style-name="GFError2"/);assert.match(chartXml,/chart:regression-curve chart:style-name="GFTrend2"/);
 assert.deepEqual(extractScatterPairData(book,'A1:E3',true).scatterXColumns,[1,3]);
 doc.active.tools.hiddenColumns=[1];const hiddenX=extractScatterPairData(book,'A1:E3',false);assert.deepEqual(hiddenX.scatterXColumns,[3]);assert.deepEqual(hiddenX.scatterPairIndexes,[1]);assert.deepEqual(hiddenX.series.map(item=>item.name),['West']);assert.match(chartSvg(book,chart),/stroke="#00aa22"/);const hiddenXXlsx=await zipRead(exportXlsx(doc)),hiddenXXml=hiddenXXlsx.get('xl/charts/chart1.xml')??'';assert.equal((hiddenXXml.match(/<c:ser>/g)??[]).length,2);assert.match(hiddenXXml,/<c:plotVisOnly val="1"\/>/);const hiddenXOds=await zipRead(exportOds(doc)),hiddenXOdsPath=[...hiddenXOds.keys()].find(path=>path.startsWith('Object ')&&path.endsWith('/content.xml')),hiddenXOdsXml=hiddenXOdsPath?hiddenXOds.get(hiddenXOdsPath)??'':'';assert.equal((hiddenXOdsXml.match(/<chart:series /g)??[]).length,1);assert.match(hiddenXOdsXml,/chart:error-indicator[^>]*chart:style-name="GFError2"/);
 doc.active.tools.hiddenColumns=[1,2,3,4];assert.throws(()=>extractScatterPairData(book,'A1:E3',false),/No visible XY column pairs/);
});
test('Scatter pair header inference tolerates numeric and blank header cells',()=>{
 const book=new Workbook(8,8);book.applyEdits([
  {ref:'A1',raw:'Label'},{ref:'B1',raw:'2024'},{ref:'C1',raw:'East'},{ref:'D1',raw:''},{ref:'E1',raw:''},
  {ref:'A2',raw:'P1'},{ref:'B2',raw:'1'},{ref:'C2',raw:'10'},{ref:'D2',raw:'2'},{ref:'E2',raw:'20'},
  {ref:'A3',raw:'P2'},{ref:'B3',raw:'3'},{ref:'C3',raw:'30'},{ref:'D3',raw:'4'},{ref:'E3',raw:'40'},
 ],false);
 const data=extractScatterPairData(book,'A1:E3');assert.equal(data.rows,2);assert.deepEqual(data.labels,['P1','P2']);
 assert.deepEqual(data.series.map(item=>item.name),['East','Series 2']);assert.deepEqual(data.scatterXValues,[[1,3],[2,4]]);
 const numericIds=new Workbook(6,8);numericIds.applyEdits([{ref:'A1',raw:'Point ID'},{ref:'B1',raw:'2024'},{ref:'C1',raw:'2025'},{ref:'D1',raw:'2030'},{ref:'E1',raw:'2040'},{ref:'A2',raw:'101'},{ref:'B2',raw:'1'},{ref:'C2',raw:'10'},{ref:'D2',raw:'2'},{ref:'E2',raw:'20'},{ref:'A3',raw:'102'},{ref:'B3',raw:'3'},{ref:'C3',raw:'30'},{ref:'D3',raw:'4'},{ref:'E3',raw:'40'}],false);const numericIdData=extractScatterPairData(numericIds,'A1:E3');assert.deepEqual(numericIdData.labels,['101','102']);assert.equal(numericIdData.rows,2);assert.deepEqual(numericIdData.series.map(item=>item.name),['2025','2040']);
});
test('XY scatter charts support 24 independent native series',async()=>{
 const book=new Workbook(8,60),edits=[];for(let pair=0;pair<24;pair++){const xCol=pair*2,yCol=xCol+1;edits.push({ref:`${colToName(xCol)}1`,raw:`X ${pair+1}`},{ref:`${colToName(yCol)}1`,raw:`Series ${pair+1}`});for(let row=2;row<=4;row++)edits.push({ref:`${colToName(xCol)}${row}`,raw:String(pair*100+row)},{ref:`${colToName(yCol)}${row}`,raw:String(pair*10+row)});}
 book.applyEdits(edits,false);const range=`A1:${colToName(47)}4`,chart={...base,id:'chart-scatter24',kind:'scatter',range,scatterPairs:true},doc=new SpreadsheetDocument(book);doc.active.tools.charts.push(chart);
 const data=extractScatterPairData(book,range);assert.equal(data.series.length,24);assert.equal(data.scatterXColumns?.length,24);assert.equal(data.scatterXValues?.[23]?.[0],2302);
 const bytes=exportXlsx(doc),xml=(await zipRead(bytes)).get('xl/charts/chart1.xml');assert.equal((xml.match(/<c:ser>/g)??[]).length,24);assert.match(xml,/<c:f>&apos;Sheet 1&apos;!\$AU\$2:\$AU\$4<\/c:f>/);assert.match(xml,/<c:f>&apos;Sheet 1&apos;!\$AV\$2:\$AV\$4<\/c:f>/);
 const restored=(await importXlsx(bytes)).document.active.tools.charts[0];assert.equal(restored?.scatterPairs,true);assert.equal(restored?.range,range);
 const ods=exportOds(doc),parts=await zipRead(ods),chartPath=[...parts.keys()].find(path=>path.startsWith('Object ')&&path.endsWith('/content.xml')),chartXml=chartPath?parts.get(chartPath)??'':'';assert.match(chartXml,/chart:domain table:cell-range-address="local-table\.\$AU\$2:\$AU\$4"/);assert.match(chartXml,/chart:values-cell-range-address="local-table\.\$AV\$2:\$AV\$4"/);assert.equal((chartXml.match(/<chart:series /g)??[]).length,24);
 const odsRestored=(await importOds(ods)).document.active.tools.charts[0];assert.equal(odsRestored?.scatterPairs,true);assert.equal(odsRestored?.range,range);
});

import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {Workbook} from '../dist/engine/workbook.js';
import {exportOds,importOds} from '../dist/engine/ods.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';
import {zipRead,zipWrite} from '../dist/engine/zip.js';
import {validateChart} from '../dist/engine/charts.js';
import {buildPivotSummary} from '../dist/engine/pivot.js';

globalThis.DOMParser=DOMParser;

function chartDocument(kind,options={}){
 const book=new Workbook(30,12);
 book.applyEdits([
  {ref:'A1',raw:'Category'},{ref:'B1',raw:'North'},{ref:'C1',raw:'South'},
  {ref:'A2',raw:'Alpha'},{ref:'B2',raw:'10'},{ref:'C2',raw:'15'},
  {ref:'A3',raw:'Beta'},{ref:'B3',raw:'20'},{ref:'C3',raw:'25'},
  {ref:'A4',raw:'Gamma'},{ref:'B4',raw:'30'},{ref:'C4',raw:'35'},
 ],false);
 const doc=new SpreadsheetDocument(book);
 doc.active.tools.charts.push(validateChart({id:'chart-odsRegression',kind,title:'Axis check',range:'A1:C4',x:0,y:0,width:500,height:300,...options},30,12));
 return doc;
}

async function embeddedChartXml(doc){
 const archive=await zipRead(exportOds(doc));
 const name=[...archive.keys()].find(path=>/^Object [^/]+\/content\.xml$/.test(path));
 assert.ok(name,'ODS export should contain an editable native chart object');
 const xml=archive.get(name),parsed=new DOMParser().parseFromString(xml,'application/xml');assert.equal(parsed.documentElement.localName,'document-content','Embedded chart content must be well-formed XML');assert.equal(parsed.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:chart:1.0','chart').length,1,'Embedded chart content must contain one ODF chart root');return xml;
}

test('ODS chart attaches standalone category rotation and reverse-order styles',async()=>{
 const xml=await embeddedChartXml(chartDocument('line',{xAxisLabelAngle:45,reverseCategories:true}));
 assert.match(xml,/<chart:axis chart:dimension="x" chart:name="x" chart:style-name="GFAxisX"/);
 assert.match(xml,/style:rotation-angle="45deg"/);
 assert.match(xml,/chart:reverse-direction="true"/);
});

test('ODS native column chart retains gap width and overlap without private chart metadata',async()=>{
 const doc=chartDocument('column',{categoryGapWidth:42,seriesOverlap:65}),archive=await zipRead(exportOds(doc)),xml=await embeddedChartXml(doc);assert.match(xml,/chart:gap-width="42"/);assert.match(xml,/chart:overlap="65"/);
 const content=archive.get('content.xml');assert.ok(content);archive.set('content.xml',content.replace(/GridForgeChart:[^<]*/g,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)})),restored=imported.document.active.tools.charts[0];assert.equal(restored?.categoryGapWidth,42);assert.equal(restored?.seriesOverlap,65);
});

test('ODS native line chart preserves per-series dash styles without private chart metadata',async()=>{
 const doc=chartDocument('line',{seriesOptions:[{lineStyle:'dashDot',lineWidth:3,markerSize:12},{lineStyle:'dot',lineWidth:.75,markerSize:8}]}),archive=await zipRead(exportOds(doc)),xml=await embeddedChartXml(doc);
 assert.match(xml,/svg:stroke-dasharray="6,3,1,3"/);assert.match(xml,/svg:stroke-dasharray="1,3"/);assert.match(xml,/svg:stroke-width="3pt"/);assert.match(xml,/svg:stroke-width="0.75pt"/);assert.match(xml,/chart:symbol-width="12pt"/);assert.match(xml,/chart:symbol-height="8pt"/);
 const content=archive.get('content.xml');assert.ok(content);archive.set('content.xml',content.replace(/GridForgeChart:[^<]*/g,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)})),options=imported.document.active.tools.charts[0]?.seriesOptions;
 assert.deepEqual(options?.map(item=>item.lineStyle),['dashDot','dot']);assert.deepEqual(options?.map(item=>item.markerSize),[12,8]);assert.deepEqual(options?.map(item=>item.lineWidth),[3,.75]);
});

test('ODS combo charts round-trip mixed series types, axis assignment, and axis titles',async()=>{
 const options={seriesOptions:[{kind:'column',axis:'primary',color:'#2563eb'},{kind:'line',axis:'secondary',color:'#dc5b73',marker:true}],yAxisTitle:'Revenue',secondaryYAxisTitle:'Growth'};
 const doc=chartDocument('combo',options),xml=await embeddedChartXml(doc);
 assert.match(xml,/chart:series[^>]*chart:class="chart:bar"/);
 assert.match(xml,/chart:series[^>]*chart:class="chart:line" chart:attached-axis="y2"/);
 const imported=await importOds(exportOds(doc)),restored=imported.document.active.tools.charts[0];
 assert.ok(restored,JSON.stringify(imported.warnings));
 assert.equal(restored.kind,'combo');assert.deepEqual(restored.seriesOptions,options.seriesOptions);
 assert.equal(restored.yAxisTitle,options.yAxisTitle);assert.equal(restored.secondaryYAxisTitle,options.secondaryYAxisTitle);
});
test('XLSX and ODS preserve scatter blank-cell spanning behavior',async()=>{
 const doc=chartDocument('scatter',{displayBlanksAs:'span'}),odsXml=await embeddedChartXml(doc);
 assert.match(odsXml,/chart:treat-empty-cells="ignore"/);
 const fromOds=await importOds(exportOds(doc));assert.equal(fromOds.document.active.tools.charts[0]?.displayBlanksAs,'span');
 const xlsxArchive=await zipRead(exportXlsx(doc));assert.match(xlsxArchive.get('xl/charts/chart1.xml'),/<c:dispBlanksAs val="span"\/>/);
 const fromXlsx=await importXlsx(exportXlsx(doc));assert.equal(fromXlsx.document.active.tools.charts[0]?.displayBlanksAs,'span');
});
test('ODS uses native Calc extensions for color scales and data bars',async()=>{
 const doc=chartDocument('line');doc.active.tools.conditional=[
  {start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#fff2dc',midColor:'#f7c65a',color:'#1a9958'},
  {start:'C2',end:'C4',operator:'dataBar',threshold:0,color:'#2878c8',negativeColor:'#cc4455',barStyle:'solid',showValue:false,minValue:-10,maxValue:50,barAxis:'none',barAxisColor:'#334455',barMinLength:8,barMaxLength:75},
  {start:'D2',end:'D4',operator:'scale',threshold:0,startColor:'#ffffff',color:'#0000ff'},
  {start:'E2',end:'E4',operator:'dataBar',threshold:0,color:'#4472c4'},
 ];
 const archive=await zipRead(exportOds(doc)),content=archive.get('content.xml');assert.ok(content);const parsed=new DOMParser().parseFromString(content,'application/xml');
 const extension='urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0';assert.equal(parsed.getElementsByTagNameNS(extension,'color-scale').length,2);assert.equal(parsed.getElementsByTagNameNS(extension,'data-bar').length,2);
 const target=parsed.getElementsByTagNameNS(extension,'conditional-format').item(0);assert.equal(target.getAttributeNS(extension,'target-range-address'),"$'Sheet 1'.$B$2:$'Sheet 1'.$B$4");assert.match(content,/calcext:gradient="false"/);assert.match(content,/calcext:show-value="false"/);assert.match(content,/calcext:axis-position="none"/);assert.match(content,/calcext:axis-color="#334455"/);assert.match(content,/calcext:min-length="8"/);assert.match(content,/calcext:max-length="75"/);
 // Remove GridForge's private metadata so this exercises the interoperable ODF extension importer.
 archive.set('content.xml',content.replace(/ gridforge:sheet-objects="[^"]*"/g,'').replace('xmlns:calcext=','xmlns:loext=').replace(/calcext:(?=[A-Za-z])/g,'loext:'));const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.conditional;
 assert.equal(restored.length,4);assert.deepEqual(restored[0],{start:'B2',end:'B4',operator:'scale',threshold:0,color:'#1a9958',startColor:'#fff2dc',midColor:'#f7c65a',scaleThresholds:[{type:'min'},{type:'percentile',value:50},{type:'max'}]});
 assert.deepEqual(restored[1],{start:'C2',end:'C4',operator:'dataBar',threshold:0,color:'#2878c8',negativeColor:'#cc4455',minValue:-10,maxValue:50,barStyle:'solid',showValue:false,barAxis:'none',barAxisColor:'#334455',barMinLength:8,barMaxLength:75});
 assert.deepEqual(restored[2],{start:'D2',end:'D4',operator:'scale',threshold:0,color:'#0000ff',startColor:'#ffffff',scaleThresholds:[{type:'min'},{type:'max'}]});
 assert.deepEqual(restored[3],{start:'E2',end:'E4',operator:'dataBar',threshold:0,color:'#4472c4',barStyle:'gradient',showValue:true});
});
test('ODS Calc color scales round-trip numeric, percentile, formula, and percent thresholds natively',async()=>{
 const doc=chartDocument('line'),three={start:'B2',end:'B4',operator:'scale',threshold:0,startColor:'#fff2dc',midColor:'#f7c65a',color:'#1a9958',scaleThresholds:[{type:'num',value:10},{type:'percentile',value:60},{type:'formula',value:'=B2+20'}]},two={start:'C2',end:'C4',operator:'scale',threshold:0,startColor:'#ffffff',color:'#0000ff',scaleThresholds:[{type:'percent',value:25},{type:'max'}]};doc.active.tools.conditional=[three,two];
 const archive=await zipRead(exportOds(doc)),content=archive.get('content.xml');assert.match(content,/calcext:type="value" calcext:color="#fff2dc"/);assert.match(content,/calcext:type="percentile" calcext:color="#f7c65a"/);assert.match(content,/calcext:type="formula" calcext:color="#1a9958"/);assert.match(content,/calcext:type="percent" calcext:color="#ffffff"/);
 archive.set('content.xml',content.replace(/ gridforge:sheet-objects="[^"]*"/g,''));const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.conditional;assert.deepEqual(restored,[{...three,scaleThresholds:[{type:'num',value:10},{type:'percentile',value:60},{type:'formula',value:'=(B2+20)'}]},two]);
});
test('ODS Calc icon sets round-trip native names, thresholds and value visibility',async()=>{
 const doc=chartDocument('line'),rule={start:'B2',end:'B4',operator:'iconSet',iconSet:'3Arrows',threshold:0,color:'#000000',iconThresholds:[{type:'min'},{type:'percentile',value:40},{type:'formula',value:'=B2+20'}],iconsOnly:true,reverseIcons:true};doc.active.tools.conditional.push(rule);const archive=await zipRead(exportOds(doc)),content=archive.get('content.xml');assert.match(content,/calcext:icon-set-type="3Arrows" calcext:show-value="false" calcext:reverse="true"/);assert.match(content,/calcext:type="percentile" calcext:value="40"/);assert.match(content,/calcext:type="formula" calcext:value="of:=\(\[\.B2\]\+20\)"/);
 archive.set('content.xml',content.replace(/ gridforge:sheet-objects="[^"]*"/g,''));const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.conditional;assert.deepEqual(restored,[{start:'B2',end:'B4',operator:'iconSet',threshold:0,color:'#000000',iconSet:'3Arrows',iconThresholds:[{type:'min'},{type:'percentile',value:40},{type:'formula',value:'=(B2+20)'}],iconsOnly:true,reverseIcons:true}]);
});
test('ODS exports and imports native formula conditional-style maps',async()=>{
 const doc=chartDocument('line'),rule={start:'B2',end:'B4',operator:'formula',formula:'=AND(B2>0,A2<>"")',threshold:0,color:'#22aa44'};doc.active.tools.conditional.push(rule);
 const bytes=exportOds(doc),archive=await zipRead(bytes),content=archive.get('content.xml');assert.match(content,/style:condition="of:is-true-formula\(/);const restored=(await importOds(bytes)).document.active.tools.conditional;assert.deepEqual(restored,[rule]);
 archive.set('content.xml',content.replace(/ gridforge:sheet-objects="[^"]*"/g,''));const nativeOnly=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.conditional;assert.equal(nativeOnly.length,1);assert.equal(nativeOnly[0].operator,'formula');assert.equal(nativeOnly[0].formula,'=AND((B2>0),(A2<>""))');assert.equal(nativeOnly[0].start,'B2');assert.equal(nativeOnly[0].end,'B4');assert.equal(nativeOnly[0].color,'#22aa44');
});
test('ODS charts retain 24 linked series and their source range',async()=>{
 const book=new Workbook(40,32),edits=[];
 for(let col=0;col<=24;col++){const letter=String.fromCharCode(65+col);edits.push({ref:`${letter}1`,raw:col===0?'Category':`Series ${col}`});}
 for(let row=2;row<=21;row++){edits.push({ref:`A${row}`,raw:`Category ${row-1}`});for(let col=1;col<=24;col++)edits.push({ref:`${String.fromCharCode(65+col)}${row}`,raw:String(row*col)});}
 book.applyEdits(edits,false);const doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-many-series',kind:'line',title:'Many series',range:'A1:Y21',x:0,y:0,width:500,height:300});
 const restored=(await importOds(exportOds(doc))).document.active.tools.charts[0];assert.ok(restored);assert.equal(restored.range,'A1:Y21');
});
test('ODS pie chart preserves more than 250 per-slice colors',async()=>{
 const book=new Workbook(400,4),edits=[{ref:'A1',raw:'Category'},{ref:'B1',raw:'Value'}],colors=[];
 for(let row=2;row<=301;row++){edits.push({ref:`A${row}`,raw:`Slice ${row-1}`},{ref:`B${row}`,raw:String(row)});colors.push(row%2?'#2563eb':'#dc5b73');}
 book.applyEdits(edits,false);const doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-colored-pie',kind:'pie',title:'Colored slices',range:'A1:B301',x:0,y:0,width:500,height:300,dataPointColors:colors});
 const xml=await embeddedChartXml(doc);assert.equal((xml.match(/<chart:data-point /g)??[]).length,300);
 const restored=(await importOds(exportOds(doc))).document.active.tools.charts[0];assert.equal(restored.dataPointColors.length,300);assert.deepEqual(restored.dataPointColors,colors);
});
test('ODS column charts preserve category-specific colors through native chart import',async()=>{
 const book=new Workbook(20,8);book.applyEdits([{ref:'A1',raw:'Category'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'North'},{ref:'B2',raw:'10'},{ref:'A3',raw:'South'},{ref:'B3',raw:'20'},{ref:'A4',raw:'West'},{ref:'B4',raw:'30'}],false);
 const colors=['#2563eb','#dc5b73','#18865b'],doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-category-colors',kind:'column',title:'Category colors',range:'A1:B4',x:0,y:0,width:500,height:300,dataPointColors:colors});
 const xml=await embeddedChartXml(doc);assert.equal((xml.match(/<chart:data-point /g)??[]).length,3);
 const archive=await zipRead(exportOds(doc));archive.set('content.xml',archive.get('content.xml').replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);
});
test('ODS combo chart colors columns and line markers and survives metadata-free import',async()=>{
 const book=new Workbook(20,8);book.applyEdits([{ref:'A1',raw:'Category'},{ref:'B1',raw:'Sales'},{ref:'C1',raw:'Margin'},{ref:'A2',raw:'North'},{ref:'B2',raw:'10'},{ref:'C2',raw:'0.1'},{ref:'A3',raw:'South'},{ref:'B3',raw:'20'},{ref:'C3',raw:'0.2'},{ref:'A4',raw:'West'},{ref:'B4',raw:'30'},{ref:'C4',raw:'0.3'}],false);
 const colors=['#2563eb','#dc5b73','#18865b'],doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-combo-colors',kind:'combo',title:'Combo colors',range:'A1:C4',x:0,y:0,width:500,height:300,dataPointColors:colors,seriesOptions:[{kind:'column'},{kind:'line'}]});
 const xml=await embeddedChartXml(doc);assert.equal((xml.match(/<chart:data-point /g)??[]).length,6);
 const archive=await zipRead(exportOds(doc));archive.set('content.xml',archive.get('content.xml').replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);
});
test('ODS line chart marker colors survive native metadata-free import',async()=>{
 const book=new Workbook(20,8);book.applyEdits([{ref:'A1',raw:'Quarter'},{ref:'B1',raw:'Revenue'},{ref:'A2',raw:'Q1'},{ref:'B2',raw:'10'},{ref:'A3',raw:'Q2'},{ref:'B3',raw:'20'},{ref:'A4',raw:'Q3'},{ref:'B4',raw:'30'}],false);
 const colors=['#2563eb','#dc5b73','#18865b'],doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-line-marker-colors',kind:'line',title:'Line marker colors',range:'A1:B4',x:0,y:0,width:500,height:300,dataPointColors:colors,seriesOptions:[{marker:true,markerStyle:'diamond'}]});
 const xml=await embeddedChartXml(doc);assert.equal((xml.match(/<chart:data-point /g)??[]).length,3);
 const archive=await zipRead(exportOds(doc));archive.set('content.xml',archive.get('content.xml').replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.charts[0];assert.deepEqual(restored.dataPointColors,colors);
});
test('ODS line, scatter, and radar charts retain independent point-color arrays per series',async()=>{
 const book=new Workbook(20,8);book.applyEdits([{ref:'A1',raw:'Quarter'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Cost'},{ref:'A2',raw:'Q1'},{ref:'B2',raw:'10'},{ref:'C2',raw:'4'},{ref:'A3',raw:'Q2'},{ref:'B3',raw:'20'},{ref:'C3',raw:'8'},{ref:'A4',raw:'Q3'},{ref:'B4',raw:'30'},{ref:'C4',raw:'12'}],false);
 const first=['#2563eb','#dc5b73','#18865b'],second=['#e04040','#40a080','#4040e0'],scatterBook=new Workbook(20,8);scatterBook.applyEdits([{ref:'A1',raw:'X'},{ref:'B1',raw:'Y1'},{ref:'C1',raw:'Y2'},{ref:'A2',raw:'1'},{ref:'B2',raw:'10'},{ref:'C2',raw:'5'},{ref:'A3',raw:'2'},{ref:'B3',raw:'20'},{ref:'C3',raw:'15'},{ref:'A4',raw:'3'},{ref:'B4',raw:'30'},{ref:'C4',raw:'25'}],false);for(const kind of ['line','radar','scatter']){const chartBook=kind==='scatter'?scatterBook:book,doc=new SpreadsheetDocument(chartBook);doc.active.tools.charts.push({id:`chart-ods-${kind}-series-colors`,kind,title:'Series colors',range:'A1:C4',x:0,y:0,width:500,height:300,seriesOptions:[{dataPointColors:first},{dataPointColors:second}]});
 const archive=await zipRead(exportOds(doc));archive.set('content.xml',archive.get('content.xml').replace(/ gridforge:sheet-objects="[^"]*"/,''));const restored=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.tools.charts[0];assert.deepEqual(restored.seriesOptions?.[0]?.dataPointColors,first,kind);assert.deepEqual(restored.seriesOptions?.[1]?.dataPointColors,second,kind);}
});
test('ODS retains the full 10,000-point single-series chart capacity',async()=>{
 const book=new Workbook(10_100,4),edits=[{ref:'A1',raw:'Category'},{ref:'B1',raw:'Value'}];
 for(let row=2;row<=10_001;row++)edits.push({ref:`A${row}`,raw:`Point ${row-1}`},{ref:`B${row}`,raw:String(row)});
 book.applyEdits(edits,false);const doc=new SpreadsheetDocument(book);doc.active.tools.charts.push({id:'chart-ods-10000-points',kind:'line',title:'Long series',range:'A1:B10001',x:0,y:0,width:500,height:300});
 const xml=await embeddedChartXml(doc);assert.match(xml,/table:cell-range-address="local-table\.\$A\$1:\$B\$10001"/);
 const restored=(await importOds(exportOds(doc))).document.active.tools.charts[0];assert.ok(restored);assert.equal(restored.range,'A1:B10001');
});

test('workbook, XLSX, and ODS preserve sheets above the former 12-sheet ceiling',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([{ref:'A1',raw:'first'}],false);
 for(let index=2;index<=16;index++){const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:`sheet-${index}`}],false);doc.addSheet(`Sheet ${index}`,book);}
 const restoredJson=SpreadsheetDocument.fromJSON(doc.toJSON()),imported=await importOds(exportOds(doc)),xlsxImported=await importXlsx(exportXlsx(doc));assert.equal(restoredJson.sheets.length,16);assert.equal(imported.document.sheets.length,16);assert.equal(imported.document.sheets[15].book.getValue('A1').value,'sheet-16');assert.equal(xlsxImported.document.sheets.length,16);assert.equal(xlsxImported.document.sheets[15].book.getValue('A1').value,'sheet-16');
});

test('sparse workbook reaches the Calc/Excel XFD1048576 address in JSON, XLSX, and ODS',async()=>{
 const book=new Workbook(1_048_576,16_384);book.applyEdits([{ref:'XFD1048576',raw:'42'}],false);const doc=new SpreadsheetDocument(book);
 const fromJson=SpreadsheetDocument.fromJSON(doc.toJSON()),fromXlsx=await importXlsx(exportXlsx(doc)),fromOds=await importOds(exportOds(doc));
 for(const restored of [fromJson,fromXlsx.document,fromOds.document]){assert.equal(restored.active.book.rows,1_048_576);assert.equal(restored.active.book.cols,16_384);assert.deepEqual(restored.active.book.getValue('XFD1048576'),{value:42});}
});

test('sparse formulas aggregate full-height and full-width ranges without dense grid allocation',()=>{
 const book=new Workbook(1_048_576,16_384);book.applyEdits([{ref:'A1',raw:'5'},{ref:'A1048576',raw:'7'},{ref:'A3',raw:'5'},{ref:'XFD3',raw:'11'},{ref:'B1',raw:'=SUM(A1:A1048576)'},{ref:'B2',raw:'=SUM(A3:XFD3)'}],false);
 assert.deepEqual(book.getValue('B1'),{value:17});assert.deepEqual(book.getValue('B2'),{value:16});
});

test('ODS chart attaches logarithmic and custom-number-format value-axis styles without bounds',async()=>{
 const xml=await embeddedChartXml(chartDocument('line',{yAxisLogBase:10,yAxisFormatCode:'$#,##0.00'}));
 assert.match(xml,/<chart:axis chart:dimension="y" chart:name="y" chart:style-name="GFAxisY"/);
 assert.match(xml,/chart:logarithmic="true"/);
 assert.match(xml,/chart:link-data-style-to-source="true"/);
});

test('ODS horizontal bar chart places category rotation and reversal on Y and numeric formatting on X',async()=>{
 const doc=chartDocument('bar',{xAxisLabelAngle:30,reverseCategories:true,yAxisMin:0,yAxisMax:50,yAxisFormatCode:'0.0%'});
 const xml=await embeddedChartXml(doc);
 assert.match(xml,/<chart:axis chart:dimension="y" chart:name="y" chart:style-name="GFAxisY"/);
 assert.match(xml,/<chart:axis chart:dimension="x" chart:name="x" chart:style-name="GFAxisX"/);
 assert.match(xml,/style:rotation-angle="30deg"/);
 assert.match(xml,/chart:reverse-direction="true"/);
 assert.match(xml,/chart:minimum="0"/);
 assert.match(xml,/chart:maximum="50"/);
 const roundTrip=await importOds(exportOds(doc)),restored=roundTrip.document.sheets[0].tools.charts[0];
 assert.equal(restored.kind,'bar');
 assert.equal(restored.xAxisLabelAngle,30);
 assert.equal(restored.reverseCategories,true);
 assert.equal(restored.yAxisMin,0);
 assert.equal(restored.yAxisMax,50);
 assert.equal(restored.yAxisFormatCode,'0.0%');
});

test('ODS exports editable native column-chart data for Histogram bins',async()=>{
 const doc=chartDocument('histogram',{histogramBins:4}),xml=await embeddedChartXml(doc);assert.match(xml,/chart:class="chart:bar" chart:vertical="true"/);assert.match(xml,/<text:p>Frequency<\/text:p>/);assert.match(xml,/table:table-row>/);assert.match(xml,/10–16\.25/);
 const imported=await importOds(exportOds(doc));assert.equal(imported.document.sheets[0].tools.charts[0].kind,'histogram');
});

test('ODS exports editable native bar and cumulative-line data for Pareto charts',async()=>{
 const doc=chartDocument('pareto'),xml=await embeddedChartXml(doc);assert.match(xml,/chart:class="chart:bar" chart:vertical="true"/);assert.match(xml,/chart:class="chart:line" chart:attached-axis="y2"/);assert.match(xml,/number:percentage-style/);assert.match(xml,/Cumulative %/);
 const imported=await importOds(exportOds(doc));assert.equal(imported.document.sheets[0].tools.charts[0].kind,'pareto');
});

test('ODS exports Waterfall as an editable stacked-column chart with totals and increase/decrease series',async()=>{
 const doc=chartDocument('waterfall',{waterfallTotalPoints:[2]}),xml=await embeddedChartXml(doc);assert.match(xml,/chart:class="chart:bar" chart:vertical="true" chart:stacked="true"/);assert.equal((xml.match(/<chart:series /g)??[]).length,4);assert.match(xml,/<text:p>Base<\/text:p>/);assert.match(xml,/<text:p>Increase<\/text:p>/);assert.match(xml,/<text:p>Decrease<\/text:p>/);assert.match(xml,/<text:p>Total<\/text:p>/);assert.match(xml,/draw:fill="none"[^>]*draw:stroke="none"/);
 const imported=await importOds(exportOds(doc));assert.equal(imported.document.sheets[0].tools.charts[0].kind,'waterfall');assert.deepEqual(imported.document.sheets[0].tools.charts[0].waterfallTotalPoints,[2]);
});

test('ODS round-trips native power regression curves and equation display flags',async()=>{
 const trendline={type:'power',showEquation:true,showR2:true};
 const doc=chartDocument('scatter',{seriesOptions:[{trendline}]}),xml=await embeddedChartXml(doc);
 assert.match(xml,/chart:regression-type="power"/);
 assert.match(xml,/chart:display-equation="true" chart:display-r-square="true"/);
 const imported=await importOds(exportOds(doc));
 assert.deepEqual(imported.document.sheets[0].tools.charts[0].seriesOptions[0].trendline,trendline);
});

for(const trendline of [
 {type:'linear',showEquation:true,showR2:false},
 {type:'exponential',showEquation:false,showR2:true},
 {type:'logarithmic'},
 {type:'polynomial',order:4,showEquation:true,showR2:true},
 {type:'movingAverage',period:3},
])test(`ODS round-trips ${trendline.type} trendline configuration`,async()=>{
 const doc=chartDocument('scatter',{seriesOptions:[{trendline}]});
 const restored=(await importOds(exportOds(doc))).document.active.tools.charts[0];
 assert.ok(restored);assert.deepEqual(restored.seriesOptions?.[0]?.trendline,trendline);
});

test('ODS formula import strips Calc function namespace but retains literals and sheet identifiers',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));
 const other=new Workbook(20,10);other.applyEdits([{ref:'A1',raw:'42'}],false);doc.addSheet('COM.MICROSOFT.Sales',other);
 doc.selectSheet(doc.sheets[0].id);doc.active.book.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'},
  {ref:'B1',raw:'10'},{ref:'B2',raw:'20'},{ref:'B3',raw:'30'},
  {ref:'D1',raw:'=FORECAST.LINEAR(4,B1:B3,A1:A3)'},
  {ref:'D2',raw:'="COM.MICROSOFT.XLOOKUP("'},
  {ref:'D3',raw:"='COM.MICROSOFT.Sales'!A1"},
 ],false);
 const imported=await importOds(exportOds(doc));
 const restored=imported.document.sheets.find(sheet=>sheet.name==='Sheet 1').book;
 assert.equal(restored.getRaw('D1'),'=FORECAST.LINEAR(4,B1:B3,A1:A3)');
 assert.equal(restored.getRaw('D2'),'="COM.MICROSOFT.XLOOKUP("');
 assert.equal(restored.getRaw('D3'),'=COM.MICROSOFT.Sales!A1');
 assert.deepEqual(restored.getValue('D3'),{value:42});
});

test('ODS formula import restores mapped Calc names and keeps cached formula results',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));
 doc.active.book.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'3'},{ref:'B1',raw:'10'},{ref:'B2',raw:'20'},{ref:'B3',raw:'30'},{ref:'D1',raw:'=FORECAST.LINEAR(4,B1:B3,A1:A3)'}],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes);
 assert.match(archive.get('content.xml'),/COM\.MICROSOFT\.FORECAST\.LINEAR\(/);
 const imported=await importOds(bytes);
 assert.equal(imported.document.active.book.getRaw('D1'),'=FORECAST.LINEAR(4,B1:B3,A1:A3)');
 assert.deepEqual(imported.document.active.book.getValue('D1'),{value:40});
});
test('ODS OpenFormula import translates separators without changing semicolons inside text literals',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([
  {ref:'A1',raw:'1'},{ref:'B1',raw:'10'},{ref:'B2',raw:'20'},{ref:'B3',raw:'30'},{ref:'C1',raw:'=SUM(B1:B3)'},
 ],false);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 archive.set('content.xml',xml.replace(/table:formula="[^"]*"/,'table:formula="of:=IF([.A1]&gt;0;SUM([.B1:.B3]);&quot;x;y&quot;)"'));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}));
 assert.equal(imported.document.active.book.getRaw('C1'),'=IF(A1>0,SUM(B1:B3),"x;y")');
 assert.deepEqual(imported.document.active.book.getValue('C1'),{value:60});
});
test('ODS formulas outside the translator keep cached values and report the fallback',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'B1',raw:'=A1+A2'}],false);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');assert.match(xml,/table:formula="of:/);
 archive.set('content.xml',xml.replace(/table:formula="[^"]*"/,'table:formula="of:=[.A1]~[.A2]"'));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}));
 assert.deepEqual(imported.document.active.book.getValue('B1'),{value:3});assert.ok(imported.warnings.some(warning=>warning.includes('syntax GridForge could not translate')));
});

test('ODS exports LET through Calc Microsoft compatibility namespace and restores its formula',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([{ref:'A1',raw:'=LET(x,2,x+2)'}],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');assert.match(xml,/COM\.MICROSOFT\.LET\(/);
 const imported=await importOds(bytes);assert.match(imported.document.active.book.getRaw('A1'),/^=LET\(X,2,/);assert.equal(imported.document.active.book.getValue('A1').value,4);
});

test('ODS maps LAMBDA and ISOMITTED to Calc compatibility namespaces and restores both formulas',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([
  {ref:'A1',raw:'=LAMBDA(x,x+1)(2)'},
  {ref:'A2',raw:'=LAMBDA(x,ISOMITTED(x))()'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.LAMBDA\(/);assert.match(xml,/COM\.MICROSOFT\.ISOMITTED\(/);
 const imported=await importOds(bytes);assert.equal(imported.document.active.book.getValue('A1').value,3);assert.equal(imported.document.active.book.getValue('A2').value,true);
});

test('ODS stores cached formula errors as string-valued formula results without converting matching text',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([
  {ref:'A1',raw:'=1/0'},
  {ref:'A2',raw:"'#DIV/0!"},
  {ref:'A3',raw:'=#GETTING_DATA'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/table:formula="of:=\(1\/0\)" office:value-type="string" office:string-value="#DIV\/0!"/);
 assert.match(xml,/office:value-type="string"><text:p>#DIV\/0!<\/text:p>/);
 assert.match(xml,/office:string-value="#GETTING_DATA"/);
 const imported=await importOds(bytes);
 assert.equal(imported.document.active.book.getRaw('A1'),'=(1/0)');
 assert.equal(imported.document.active.book.getValue('A1').error.code,'#DIV/0!');
 assert.equal(imported.document.active.book.getValue('A2').value,'#DIV/0!');
 assert.equal(imported.document.active.book.getValue('A3').error.code,'#GETTING_DATA');
});

test('ODS emits known _xlfn dynamic functions through the Microsoft compatibility namespace',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'B1',raw:'1'},{ref:'B2',raw:'0'},{ref:'D1',raw:'=_xlfn.FILTER(A1:A2,B1:B2)'}],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes);assert.match(archive.get('content.xml'),/COM\.MICROSOFT\.FILTER\(/);const imported=await importOds(bytes);assert.equal(imported.document.active.book.getRaw('D1'),'=FILTER(A1:A2,B1:B2)');
});
test('ODS maps TEXTBEFORE and TEXTAFTER to Calc Microsoft functions and restores optional fallback arguments',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'=TEXTBEFORE("Socrates"," ",,,1)'},
  {ref:'A2',raw:'=TEXTAFTER("Socrates"," ",,,1)'},
  {ref:'A3',raw:'=TEXTBEFORE("abc","z",,,0,)'},
  {ref:'A4',raw:'=TEXTAFTER("abc","z",,,0,)'},
  {ref:'A5',raw:'=TEXTBEFORE("abc","z",,,0,"")'},
  {ref:'A6',raw:'=TEXTAFTER("abc","z",,,0,"")'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.TEXTBEFORE\(/);assert.match(xml,/COM\.MICROSOFT\.TEXTAFTER\(/);
 const book=(await importOds(bytes)).document.active.book;
 assert.equal(book.getRaw('A1'),'=TEXTBEFORE("Socrates"," ",,,1)');assert.equal(book.getRaw('A2'),'=TEXTAFTER("Socrates"," ",,,1)');
 assert.equal(book.getValue('A1').value,'Socrates');assert.equal(book.getValue('A2').value,'');
 assert.equal(book.getValue('A3').error?.code,'#N/A');assert.equal(book.getValue('A4').error?.code,'#N/A');
 assert.equal(book.getValue('A5').value,'');assert.equal(book.getValue('A6').value,'');
});
test('ODS array formulas preserve padded spill errors as error cells and recalculate them on import',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'1'},{ref:'C1',raw:"'x"},{ref:'C2',raw:"'y"},{ref:'E1',raw:'=VSTACK(A1:A1,C1:D2)'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.VSTACK\(/);assert.match(xml,/table:number-matrix-rows-spanned="3" table:number-matrix-columns-spanned="2"/);
 assert.match(xml,/office:value-type="error" office:string-value="#N\/A"/);
 const imported=await importOds(bytes),book=imported.document.active.book;assert.equal(book.getValue('E1').value,1);assert.equal(book.getValue('F1').error?.code,'#N/A');assert.equal(book.getValue('E2').value,'x');assert.equal(book.getValue('E3').value,'y');
});
test('ODS TEXTSPLIT preserves Excel default #N/A padding with an explicitly omitted pad slot',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'D1',raw:'=TEXTSPLIT("a,b;c",",",";",FALSE,0,)'}
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.TEXTSPLIT\(/);assert.match(xml,/office:value-type="error" office:string-value="#N\/A"/);
 const book=(await importOds(bytes)).document.active.book;assert.equal(book.getRaw('D1'),'=TEXTSPLIT("a,b;c",",",";",FALSE(),0,)');
 assert.deepEqual(['D1','E1','D2'].map(ref=>book.getValue(ref).value),['a','b','c']);assert.equal(book.getValue('E2').error?.code,'#N/A');
});
test('ODS FILTER formulas preserve selected source errors through namespace export and recalculation',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'10'},{ref:'A2',raw:'=1/0'},{ref:'A3',raw:'30'},{ref:'A4',raw:'40'},
  {ref:'B1',raw:'=TRUE()'},{ref:'B2',raw:'=TRUE()'},{ref:'B3',raw:'=FALSE()'},{ref:'B4',raw:'=TRUE()'},
  {ref:'D1',raw:'=FILTER(A1:A4,B1:B4)'},{ref:'F1',raw:'=FILTER(A1:A4,B1:B4=7,)'},{ref:'H1',raw:'=FILTER(A1:A4,B1:B4=7,"")'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.FILTER\(/);assert.match(xml,/office:value-type="error" office:string-value="#DIV\/0!"/);
 const book=(await importOds(bytes)).document.active.book;assert.equal(book.getRaw('D1'),'=FILTER(A1:A4,B1:B4)');
 assert.deepEqual(['D1','D2','D3'].map(ref=>{const cell=book.getValue(ref);return cell.error?.code??cell.value;}),[10,'#DIV/0!',40]);
 assert.equal(book.getRaw('F1'),'=FILTER(A1:A4,(B1:B4=7),)');assert.equal(book.getValue('F1').error?.code,'#CALC!');
 assert.equal(book.getRaw('H1'),'=FILTER(A1:A4,(B1:B4=7),"")');assert.equal(book.getValue('H1').value,'');
});
test('ODS round-trips spilled IFS and SWITCH formulas with recalculated cell results',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'A'},{ref:'A2',raw:'B'},{ref:'A3',raw:'C'},
  {ref:'C1',raw:'=IFS(A1:A3="A","yes",A1:A3="B","maybe",TRUE,"no")'},
  {ref:'F1',raw:'=SWITCH(A1:A3,"A",1,"B",2,99)'},
 ],false);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 assert.match(xml,/COM\.MICROSOFT\.IFS\(/);assert.match(xml,/COM\.MICROSOFT\.SWITCH\(/);
 const book=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)}))).document.active.book;
 assert.equal(book.getRaw('C1'),'=IFS((A1:A3="A"),"yes",(A1:A3="B"),"maybe",TRUE(),"no")');
 assert.equal(book.getRaw('F1'),'=SWITCH(A1:A3,"A",1,"B",2,99)');
 assert.deepEqual(['C1','C2','C3'].map(ref=>book.getValue(ref).value),['yes','maybe','no']);
 assert.deepEqual(['F1','F2','F3'].map(ref=>book.getValue(ref).value),[1,2,99]);
});
test('XLSX round-trips spilled IFS and SWITCH formulas with recalculated cell results',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'3'},{ref:'A3',raw:'-1'},
  {ref:'C1',raw:'=IFS(A1:A3>2,"high",A1:A3>0,"low",TRUE,"none")'},
  {ref:'F1',raw:'=SWITCH(A1:A3,1,"one",3,"three","other")'},
 ],false);
 const imported=await importXlsx(exportXlsx(doc)),book=imported.document.active.book;
 assert.equal(book.getRaw('C1'),'=IFS(A1:A3>2,"high",A1:A3>0,"low",TRUE,"none")');
 assert.equal(book.getRaw('F1'),'=SWITCH(A1:A3,1,"one",3,"three","other")');
 assert.deepEqual(['C1','C2','C3'].map(ref=>book.getValue(ref).value),['low','high','none']);
 assert.deepEqual(['F1','F2','F3'].map(ref=>book.getValue(ref).value),['one','three','other']);
});
test('XLSX round-trips omitted optional formula fallbacks and padding arguments',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,12));doc.active.book.applyEdits([
  {ref:'A1',raw:'1'},{ref:'A2',raw:'2'},
  {ref:'D1',raw:'=TEXTSPLIT("a,b;c",",",";",FALSE,0,)'},
  {ref:'G1',raw:'=FILTER(A1:A2,A1:A2>9,)'},
  {ref:'J1',raw:'=FILTER(A1:A2,A1:A2>9,"")'},
  {ref:'K1',raw:'=TEXTBEFORE("abc","z",,,0,)'},
  {ref:'L1',raw:'=TEXTBEFORE("abc","z",,,0,"")'}
 ],false);
 const book=(await importXlsx(exportXlsx(doc))).document.active.book;
 assert.deepEqual(['D1','E1','D2'].map(ref=>book.getValue(ref).value),['a','b','c']);assert.equal(book.getValue('E2').error?.code,'#N/A');
 assert.equal(book.getValue('G1').error?.code,'#CALC!');assert.equal(book.getValue('J1').value,'');
 assert.equal(book.getValue('K1').error?.code,'#N/A');assert.equal(book.getValue('L1').value,'');
});
test('ODS emits Calc-specific regex formulas and preserves the original GridForge formulas',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits([
  {ref:'A1',raw:'=REGEXEXTRACT("AB123","[0-9]+")'},
  {ref:'A2',raw:'=REGEXREPLACE("ab12","[0-9]+","ZZ")'},
  {ref:'A3',raw:'=REGEXTEST("ab12","[0-9]+")'},
  {ref:'A4',raw:'=REGEXEXTRACT("A12 B34","[A-Z]\\d+",1,1)'},
 ],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.equal((xml.match(/gridforge:source-formula=/g)??[]).length,4);
 assert.equal((xml.match(/table:formula=/g)??[]).length,3);
 assert.match(xml,/ORG\.LIBREOFFICE\.REGEX\(/);
 assert.match(xml,/IF\(ISNA\(ORG\.LIBREOFFICE\.REGEX\(/);
 const restored=(await importOds(bytes)).document.active.book;
 assert.deepEqual(['A1','A2','A3','A4'].map(ref=>restored.getRaw(ref)),[
  '=REGEXEXTRACT("AB123","[0-9]+")','=REGEXREPLACE("ab12","[0-9]+","ZZ")','=REGEXTEST("ab12","[0-9]+")',
  '=REGEXEXTRACT("A12 B34","[A-Z]\\d+",1,1)',
 ]);
 assert.deepEqual(['A1','A2','A3'].map(ref=>restored.getValue(ref).value),['123','abZZ',true]);
});
test('ODS imports supported LibreOffice REGEX formulas and falls back for unsupported occurrence mode',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));doc.active.book.applyEdits(['=1','=1','=1','=1'].map((raw,index)=>({ref:`A${index+1}`,raw})),false);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml'),formulas=[
  'of:=ORG.LIBREOFFICE.REGEX("AB123";"[0-9]+")',
  'of:=ORG.LIBREOFFICE.REGEX("ab12";"[0-9]+";"ZZ";"g")',
  'of:=NOT(ISERROR(ORG.LIBREOFFICE.REGEX("ab12";"[0-9]+")))',
  'of:=ORG.LIBREOFFICE.REGEX("a1 a2";"a[0-9]";"x";2)',
 ];let formulaIndex=0;
 const content=xml.replace(/table:formula="[^"]*"/g,()=>{const formula=formulas[formulaIndex++]??'';return `table:formula="${formula.replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;').replaceAll('>','&gt;')}"`;});
 assert.equal(formulaIndex,4);
 archive.set('content.xml',content);
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(archive)})),book=imported.document.active.book;
 assert.equal(book.getRaw('A1'),'=REGEXEXTRACT("AB123","[0-9]+")');
 assert.equal(book.getRaw('A2'),'=REGEXREPLACE("ab12","[0-9]+","ZZ")');
 assert.equal(book.getRaw('A3'),'=REGEXTEST("ab12","[0-9]+")');
 assert.deepEqual(['A1','A2','A3'].map(ref=>book.getValue(ref).value),['123','abZZ',true]);
 assert.deepEqual(book.getValue('A4'),{value:1});
 assert.ok(imported.warnings.some(warning=>warning.includes('1 ODS formulas used syntax GridForge could not translate')));
});

test('ODS formula export/import retains absolute cross-sheet range references',async()=>{
 const sales=new Workbook(20,10);sales.applyEdits([{ref:'A1',raw:'10'},{ref:'A2',raw:'20'},{ref:'A3',raw:'30'}],false);const doc=new SpreadsheetDocument(new Workbook(20,10));doc.addSheet('Sales Data',sales);doc.selectSheet(doc.sheets[0].id);doc.active.book.applyEdits([{ref:'D1',raw:"=SUM('Sales Data'!$A$1:$A$3)"}],false);
 const bytes=exportOds(doc),archive=await zipRead(bytes);assert.match(archive.get('content.xml'),/of:=SUM\(\[\$&apos;Sales Data&apos;\.\$A\$1:\$&apos;Sales Data&apos;\.\$A\$3\]\)/);const imported=await importOds(bytes);assert.equal(imported.document.sheets[0].book.getValue('D1').value,60);assert.equal(imported.document.sheets[0].book.getRaw('D1'),"=SUM('Sales Data'!$A$1:$A$3)");imported.document.sheets[1].book.setCell('A1','50');assert.equal(imported.document.sheets[0].book.getValue('D1').value,100);
});
test('ODS preserves XLOOKUP match and reverse-search modes while retaining GridForge results',async()=>{
 const book=new Workbook(20,12);book.applyEdits([{ref:'D1',raw:'key'},{ref:'D2',raw:'k2'},{ref:'D3',raw:'k2'},{ref:'E1',raw:'Value'},{ref:'E2',raw:'20'},{ref:'E3',raw:'30'},{ref:'A1',raw:'=XLOOKUP("k2",D2:D3,E2:E3,"missing",0,-1)'}],false);
 const doc=new SpreadsheetDocument(book),xml=(await zipRead(exportOds(doc))).get('content.xml');assert.match(xml,/COM\.MICROSOFT\.XLOOKUP\(&quot;k2&quot;;\[\.D2:\.D3\];\[\.E2:\.E3\];&quot;missing&quot;;0;-\(1\)\)/);
 const restored=(await importOds(exportOds(doc))).document.active.book;assert.equal(restored.getRaw('A1'),'=XLOOKUP("k2",D2:D3,E2:E3,"missing",0,-(1))');assert.equal(restored.getValue('A1').value,30);
});

test('ODS writes native Data Pilot fields for supported Pivot Summary rows, columns, pages, and measures',async()=>{
 const book=new Workbook(40,20);book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Quarter'},{ref:'C1',raw:'Sales'},{ref:'D1',raw:'Channel'},
  {ref:'A2',raw:'West'},{ref:'B2',raw:'Q1'},{ref:'C2',raw:'10'},{ref:'D2',raw:'Web'},
  {ref:'A3',raw:'East'},{ref:'B3',raw:'Q2'},{ref:'C3',raw:'20'},{ref:'D3',raw:'Web'},
 ],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-native',sourceRange:'A1:D3',destination:'F1',rowField:0,columnField:1,valueField:2,filterField:3,filterValue:'Web',aggregation:'sum'};
 const result=buildPivotSummary(book,spec);book.applyEdits(result.edits,false);spec.lastOutputRange=result.outputRange;doc.active.tools.pivotTables.push(spec);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 assert.match(xml,/<table:data-pilot-tables><table:data-pilot-table table:name="GF_DataPilot_1_1" table:target-range-address="\$&apos;Sheet 1&apos;\.\$F\$1" table:grand-total="both"/);
 assert.match(xml,/<table:source-cell-range table:cell-range-address="\$&apos;Sheet 1&apos;\.\$A\$1:\.\$D\$3"\/>/);
 assert.match(xml,/table:source-field-name="Region" table:orientation="row"/);
 assert.match(xml,/table:source-field-name="Quarter" table:orientation="column"/);
 assert.match(xml,/table:source-field-name="Sales" table:orientation="data" table:function="sum"/);
 assert.match(xml,/table:source-field-name="Channel" table:orientation="page" table:function="auto" table:selected-page="Web"/);
});

test('ODS imports native Data Pilot fields even when GridForge-only pivot metadata is absent',async()=>{
 const book=new Workbook(40,20);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Quarter'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'Q1'},{ref:'C2',raw:'10'},{ref:'A3',raw:'East'},{ref:'B3',raw:'Q2'},{ref:'C3',raw:'20'}],false);
 const doc=new SpreadsheetDocument(book),pivot={id:'pivot-ods-import',sourceRange:'A1:C3',destination:'E1',rowField:0,columnField:1,valueField:2,aggregation:'average'};
 const summary=buildPivotSummary(book,pivot);book.applyEdits(summary.edits,false);doc.active.tools.pivotTables.push(pivot);
 const parts=await zipRead(exportOds(doc));parts.set('content.xml',parts.get('content.xml').replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),restored=imported.document.active.tools.pivotTables[0];
 assert.ok(restored,JSON.stringify(imported.warnings));assert.equal(restored.destination,'E1');assert.equal(restored.sourceRange,'A1:C3');
 assert.deepEqual(restored.rowFields??[restored.rowField],[0]);assert.deepEqual(restored.columnFields??[restored.columnField],[1]);assert.equal(restored.aggregation,'average');
});

test('ODS Data Pilot preserves multiple measure fields with different aggregates',async()=>{
 const book=new Workbook(40,20);book.applyEdits([{ref:'A1',raw:'Team'},{ref:'B1',raw:'Sales'},{ref:'C1',raw:'Units'},{ref:'A2',raw:'A'},{ref:'B2',raw:'40'},{ref:'C2',raw:'4'},{ref:'A3',raw:'A'},{ref:'B3',raw:'60'},{ref:'C3',raw:'8'}],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-measures',sourceRange:'A1:C3',destination:'E1',rowField:0,valueField:1,valueFields:[1,2],measures:[{field:1,aggregation:'sum',caption:'Revenue'},{field:2,aggregation:'average',caption:'Average units'}],aggregation:'sum'};
 doc.active.tools.pivotTables.push(spec);const bytes=exportOds(doc),archive=await zipRead(bytes),xml=archive.get('content.xml');
 assert.match(xml,/table:source-field-name="Sales" table:orientation="data" table:function="sum"/);assert.match(xml,/table:source-field-name="Units" table:orientation="data" table:function="average"/);assert.match(xml,/table:source-field-name="Data" table:is-data-layout-field="true" table:orientation="column"/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),pivot=imported.document.active.tools.pivotTables[0];
 assert.ok(pivot,JSON.stringify(imported.warnings));assert.deepEqual(pivot.measures,[{field:1,aggregation:'sum'},{field:2,aggregation:'average'}]);
});

test('ODS omits native Data Pilot when advanced GridForge grouping cannot be represented faithfully',async()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Amount'},{ref:'A2',raw:'46000'},{ref:'B2',raw:'10'}],false);
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push({id:'pivot-ods-grouped',sourceRange:'A1:B2',destination:'D1',rowField:0,valueField:1,aggregation:'sum',dateGroups:[{field:0,unit:'weeks'}]});
 const xml=(await zipRead(exportOds(doc))).get('content.xml');assert.doesNotMatch(xml,/<table:data-pilot-tables>/);assert.match(xml,/&quot;dateGroups&quot;/);
});

test('ODS omits native Data Pilot when GridForge uses value-based field sorting',async()=>{
 const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'10'}],false);
 const doc=new SpreadsheetDocument(book);doc.active.tools.pivotTables.push({id:'pivot-ods-value-sort',sourceRange:'A1:B2',destination:'D1',rowField:0,valueField:1,aggregation:'sum',sortOrder:'valueDescending'});
 const xml=(await zipRead(exportOds(doc))).get('content.xml');assert.doesNotMatch(xml,/<table:data-pilot-tables>/);assert.match(xml,/&quot;sortOrder&quot;:&quot;valueDescending&quot;/);
});

test('ODS Data Pilot maps supported date, numeric, and manual groups to native group elements',async()=>{
 const book=new Workbook(40,20);book.applyEdits([
  {ref:'A1',raw:'Date'},{ref:'B1',raw:'Amount'},{ref:'C1',raw:'Team'},{ref:'D1',raw:'Sales'},
  {ref:'A2',raw:'45000'},{ref:'B2',raw:'12'},{ref:'C2',raw:'North'},{ref:'D2',raw:'100'},
  {ref:'A3',raw:'45030'},{ref:'B3',raw:'28'},{ref:'C3',raw:'South'},{ref:'D3',raw:'200'},
 ],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-group-types',sourceRange:'A1:D3',destination:'F1',rowField:0,rowFields:[0,1,2],valueField:3,aggregation:'sum',dateGroups:[{field:0,unit:'months'}],numberGroups:[{field:1,interval:10,start:10,end:40}],manualGroups:[{field:2,groups:[{label:'Domestic',values:['North']},{label:'International',values:['South']}]}]};
 doc.active.tools.pivotTables.push(spec);const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 assert.match(xml,/table:grouped-by="months" table:date-start="\d{4}-\d{2}-\d{2}T00:00:00" table:date-end="\d{4}-\d{2}-\d{2}T00:00:00"/);
 assert.match(xml,/table:start="10" table:end="39\.99999999999999" table:step="10"/);
 assert.match(xml,/<table:data-pilot-group table:name="Domestic"><table:data-pilot-group-member table:name="North"\/><\/table:data-pilot-group>/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),pivot=imported.document.active.tools.pivotTables[0];
 assert.ok(pivot,JSON.stringify(imported.warnings));assert.deepEqual(pivot.dateGroups,[{field:0,unit:'months'}]);assert.deepEqual(pivot.numberGroups,[{field:1,start:10,end:40,interval:10}]);
 assert.deepEqual(pivot.manualGroups,[{field:2,groups:[{label:'Domestic',values:['North']},{label:'International',values:['South']}]}]);
});

test('ODS native Data Pilot retains a cross-sheet source independently of its destination sheet',async()=>{
 const sourceBook=new Workbook(30,12);sourceBook.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'10'},{ref:'A3',raw:'East'},{ref:'B3',raw:'20'}],false);
 const doc=new SpreadsheetDocument(sourceBook),target=doc.addSheet('Pivot Report',new Workbook(30,12));doc.selectSheet(target.id);
 const spec={id:'pivot-ods-cross-sheet',sourceRange:'A1:B3',sourceSheet:doc.sheets[0].name,destination:'D1',rowField:0,valueField:1,aggregation:'sum'},summary=buildPivotSummary(target.book,spec,sourceBook);target.book.applyEdits(summary.edits,false);spec.lastOutputRange=summary.outputRange;target.tools.pivotTables.push(spec);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 assert.match(xml,/table:target-range-address="\$&apos;Pivot Report&apos;\.\$D\$1"/);assert.match(xml,/table:cell-range-address="\$&apos;Sheet 1&apos;\.\$A\$1:\.\$B\$3"/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),restored=imported.document.sheets.find(sheet=>sheet.name==='Pivot Report')?.tools.pivotTables[0];
 assert.ok(restored,JSON.stringify(imported.warnings));assert.equal(restored.sourceSheet,'Sheet 1');assert.equal(restored.destination,'D1');assert.equal(restored.sourceRange,'A1:B3');
});

test('ODS native Data Pilot round-trips multiple report-filter fields',async()=>{
 const book=new Workbook(40,20);book.applyEdits([
  {ref:'A1',raw:'Region'},{ref:'B1',raw:'Channel'},{ref:'C1',raw:'Department'},{ref:'D1',raw:'Sales'},
  {ref:'A2',raw:'West'},{ref:'B2',raw:'Online'},{ref:'C2',raw:'Retail'},{ref:'D2',raw:'30'},
  {ref:'A3',raw:'East'},{ref:'B3',raw:'Online'},{ref:'C3',raw:'Retail'},{ref:'D3',raw:'20'},
 ],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-multi-page',sourceRange:'A1:D3',destination:'F1',rowField:0,valueField:3,aggregation:'sum',filters:[{field:1,value:'Online'},{field:2,value:'Retail'}]};
 doc.active.tools.pivotTables.push(spec);const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');
 assert.match(xml,/table:source-field-name="Channel" table:orientation="page" table:function="auto" table:selected-page="Online"/);
 assert.match(xml,/table:source-field-name="Department" table:orientation="page" table:function="auto" table:selected-page="Retail"/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),pivot=imported.document.active.tools.pivotTables[0];
 assert.ok(pivot,JSON.stringify(imported.warnings));assert.deepEqual(pivot.filters,[{field:1,value:'Online'},{field:2,value:'Retail'}]);
});

test('ODS native Data Pilot round-trips selected row subtotals',async()=>{
 const book=new Workbook(30,12);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Team'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'A'},{ref:'C2',raw:'10'},{ref:'A3',raw:'West'},{ref:'B3',raw:'B'},{ref:'C3',raw:'20'}],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-partial-subtotal',sourceRange:'A1:C3',destination:'E1',rowField:0,rowFields:[0,1],valueField:2,aggregation:'sum',subtotals:true,subtotalFields:[0]};doc.active.tools.pivotTables.push(spec);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');assert.match(xml,/table:source-field-name="Region" table:orientation="row" table:function="auto"><table:data-pilot-level><table:data-pilot-sort-info[^>]*\/><table:data-pilot-subtotals><table:data-pilot-subtotal table:function="auto"/);assert.match(xml,/table:source-field-name="Team" table:orientation="row" table:function="auto"><table:data-pilot-level><table:data-pilot-sort-info[^>]*\/><\/table:data-pilot-level><\/table:data-pilot-field>/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));
 const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),pivot=imported.document.active.tools.pivotTables[0];assert.ok(pivot,JSON.stringify(imported.warnings));assert.equal(pivot.subtotals,true);assert.deepEqual(pivot.subtotalFields,[0]);
});

test('ODS native Data Pilot preserves per-field ascending and descending sort order',async()=>{
 const book=new Workbook(30,12);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Team'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'Zeta'},{ref:'C2',raw:'10'},{ref:'A3',raw:'East'},{ref:'B3',raw:'Alpha'},{ref:'C3',raw:'20'}],false);
 const doc=new SpreadsheetDocument(book),spec={id:'pivot-ods-field-sort',sourceRange:'A1:C3',destination:'E1',rowField:0,rowFields:[0,1],valueField:2,aggregation:'sum',sortOrder:'ascending',fieldSortOrders:{'0':'descending','1':'ascending'}};doc.active.tools.pivotTables.push(spec);
 const archive=await zipRead(exportOds(doc)),xml=archive.get('content.xml');assert.match(xml,/<table:data-pilot-sort-info table:sort-mode="name" table:order="descending"\/>/);assert.match(xml,/<table:data-pilot-sort-info table:sort-mode="name" table:order="ascending"\/>/);
 const parts=new Map(archive);parts.set('content.xml',xml.replace(/ gridforge:sheet-objects="[^"]*"/,''));const imported=await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(parts)})),pivot=imported.document.active.tools.pivotTables[0];assert.ok(pivot,JSON.stringify(imported.warnings));assert.deepEqual(pivot.fieldSortOrders,{'0':'descending','1':'ascending'});
});

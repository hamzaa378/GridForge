import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {fitTrendline} from '../dist/engine/trendlines.js';
import {Workbook} from '../dist/engine/workbook.js';
import {dataBarFill,ruleColor,conditionalIcon,nextConditionalPriority,validateTools,freshTools} from '../dist/engine/sheettools.js';
import {validateChart,chartSvg,calculateAxisScales} from '../dist/engine/charts.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx,importXlsx} from '../dist/engine/xlsx.js';
import {exportOds,importOds} from '../dist/engine/ods.js';
import {zipRead,zipWrite} from '../dist/engine/zip.js';

globalThis.DOMParser=DOMParser;

test('linear trendline matches independent reference slope, intercept and R squared',()=>{
 const result=fitTrendline([3,5,7,9],{type:'linear',showEquation:true,showR2:true});
 assert.deepEqual(result.points.map(p=>p.y),[3,5,7,9]);assert.equal(result.rSquared,1);assert.match(result.equation,/2\.0000/);
});
test('exponential, logarithmic, polynomial and power trendlines fit known equations',()=>{
 assert.ok(Math.abs(fitTrendline([2*Math.exp(.5),2*Math.exp(1),2*Math.exp(1.5)],{type:'exponential'}).points[0].y-2*Math.exp(.5))<1e-8);
 assert.ok(Math.abs(fitTrendline([1,1+Math.log(2),1+Math.log(3)],{type:'logarithmic'}).points[1].y-(1+Math.log(2)))<1e-8);
 assert.ok(Math.abs(fitTrendline([2,5,10,17],{type:'polynomial',order:2}).points[1].y-5)<1e-8);
 assert.ok(Math.abs(fitTrendline([2,8,18],{type:'power'}).points[1].y-8)<1e-8);
});
test('moving average and invalid trendline domains are handled',()=>{
 assert.deepEqual(fitTrendline([1,2,3,4],{type:'movingAverage',period:2}).points.map(p=>p.y),[null,1.5,2.5,3.5]);
 assert.match(fitTrendline([1,2,3],{type:'logarithmic'},[-1,2,3]).error,/positive x/);
 assert.match(fitTrendline([1,2],{type:'polynomial',order:2}).error,/lower than/);
});
test('chart series metadata validates secondary axis and persistent trendline settings',()=>{
 const chart=validateChart({id:'chart-v111',kind:'combo',title:'Sales and margin',range:'A1:C5',x:0,y:0,width:500,height:300,seriesOptions:[{}, {axis:'secondary',kind:'line',trendline:{type:'linear',forward:2,showEquation:true}}]},20,10);
 assert.equal(chart.seriesOptions[1].axis,'secondary');assert.equal(chart.seriesOptions[1].trendline.forward,2);
});
test('chart axes, gridlines, number format, series colors and markers validate and render',()=>{
 const chart=validateChart({id:'chart-axis-ui',kind:'line',title:'Axis options',range:'A1:C5',x:0,y:0,width:500,height:300,yAxisMin:0,yAxisMax:200,yAxisNumberFormat:'currency',majorGridlines:false,minorGridlines:true,seriesOptions:[{color:'#d12b52',marker:false}]},20,10);
 assert.equal(chart.yAxisMin,0);assert.equal(chart.yAxisMax,200);assert.equal(chart.yAxisNumberFormat,'currency');assert.equal(chart.majorGridlines,false);assert.equal(chart.minorGridlines,true);
 const wb=new Workbook(20,10);wb.applyEdits([{ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Margin'},{ref:'A2',raw:'Jan'},{ref:'B2',raw:'100'},{ref:'C2',raw:'20'},{ref:'A3',raw:'Feb'},{ref:'B3',raw:'200'},{ref:'C3',raw:'30'}],false);const svg=chartSvg(wb,chart);assert.match(svg,/stroke="#d12b52"/);assert.match(svg,/<rect x="58" y="280" width="10" height="10" fill="#d12b52"\/>/);assert.match(svg,/\$100/);assert.match(svg,/stroke="#f1f4f8"/);
 assert.throws(()=>validateChart({...chart,yAxisMin:501},20,10),/minimum must be less/);
});
test('series error bars render, validate and persist through XLSX chart parts',async()=>{
 const chart=validateChart({id:'chart-errors',kind:'line',title:'Uncertainty',range:'A1:B4',x:0,y:0,width:500,height:300,seriesOptions:[{errorBars:{type:'fixed',value:2}}]},12,8);
 const wb=new Workbook(12,8);wb.applyEdits([{ref:'A1',raw:'Period'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'1'},{ref:'B2',raw:'10'},{ref:'A3',raw:'2'},{ref:'B3',raw:'14'},{ref:'A4',raw:'3'},{ref:'B4',raw:'12'}],false);
 assert.match(chartSvg(wb,chart),/<g stroke="#4472c4"[^>]*><line/);
 assert.throws(()=>validateChart({...chart,seriesOptions:[{errorBars:{type:'fixed',value:-1}}]},12,8),/error bars/);
 const doc=new SpreadsheetDocument(wb);doc.active.tools.charts.push(chart);const zip=await zipRead(exportXlsx(doc)),xml=zip.get('xl/charts/chart1.xml');assert.match(xml,/<c:errValType val="fixedVal"\/><c:val val="2"\/>/);
});
test('formula conditional formatting shifts relative references, honors locks and survives XLSX roundtrip',async()=>{
 assert.equal(nextConditionalPriority([]),1);assert.equal(nextConditionalPriority([{priority:4},{priority:1}]),5);assert.equal(nextConditionalPriority([{operator:'gt'},{operator:'lt'}]),3);
 const wb=new Workbook(12,8);wb.applyEdits([{ref:'A1',raw:'Name'},{ref:'B1',raw:'Actual'},{ref:'C1',raw:'Limit'},{ref:'A2',raw:'a'},{ref:'B2',raw:'12'},{ref:'C2',raw:'10'},{ref:'A3',raw:'b'},{ref:'B3',raw:'8'},{ref:'C3',raw:'10'},{ref:'A4',raw:'c'},{ref:'B4',raw:'20'},{ref:'C4',raw:'10'},{ref:'D2',raw:'=B2*2'}],false);
 const doc=new SpreadsheetDocument(wb),rule={start:'B2',end:'B4',operator:'formula',formula:'=B2>$C2',threshold:0,priority:2,stopIfTrue:true,color:'#22aa44'};doc.active.tools.conditional.push(rule);
 const tools=validateTools({...freshTools(),conditional:[rule]},12,8);assert.equal(tools.conditional[0].formula,'=B2>$C2');
 assert.equal(ruleColor('B2',12,[rule],wb),'#22aa44');assert.equal(ruleColor('B3',8,[rule],wb),null);assert.equal(ruleColor('B4',20,[rule],wb),'#22aa44');
 wb.setCell('C2','13');assert.equal(ruleColor('B2',12,[rule],wb),null);wb.setCell('C2','11');assert.equal(ruleColor('B2',12,[rule],wb),'#22aa44');wb.setCell('C2','10');
 assert.equal(wb.evaluateFormulaAt('=B2>$C2','B2').value,true);assert.equal(wb.evaluateFormulaAt('=B3>$C3','B3').value,false);
 assert.equal(wb.evaluateFormulaAt('=ROW()=2','B2').value,true);assert.equal(ruleColor('D2',24,[{start:'D2',end:'D2',operator:'formula',formula:'=D2>20',threshold:0,color:'#ff0000'}],wb),'#ff0000');
 doc.active.tools.conditional.push({...rule,formula:'=B2<15',priority:1,stopIfTrue:false,color:'#aa2244'});const archive=await zipRead(exportXlsx(doc)),xml=archive.get('xl/worksheets/sheet1.xml');assert.match(xml,/<cfRule type="expression" priority="2" stopIfTrue="1"[^>]*><formula>B2&gt;\$C2<\/formula>/);assert.match(xml,/<cfRule type="expression" priority="1"[^>]*><formula>B2&lt;15<\/formula>/);
 const restoredRules=(await importXlsx(exportXlsx(doc))).document.active.tools.conditional;assert.equal(restoredRules.length,2);assert.equal(restoredRules[0].formula,'=B2<15');assert.equal(restoredRules[0].priority,1);assert.equal(restoredRules[0].stopIfTrue,undefined);const restored=restoredRules[1];assert.equal(restored.operator,'formula');assert.equal(restored.formula,'=B2>$C2');assert.equal(restored.color,'#22aa44');assert.equal(restored.priority,2);assert.equal(restored.stopIfTrue,true);assert.equal(ruleColor('B2',12,restoredRules,wb),'#aa2244');assert.equal(ruleColor('B4',20,[restored],wb),'#22aa44');
 const cfParts=[...xml.matchAll(/<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>/g)].map(match=>match[0]),combined=`<conditionalFormatting sqref="B2:B4">${cfParts.map(part=>part.match(/<cfRule\b[\s\S]*?<\/cfRule>/)?.[0]??'').join('')}</conditionalFormatting>`,multiRuleXml=xml.replace(cfParts.join(''),combined);archive.set('xl/worksheets/sheet1.xml',multiRuleXml);const multiRuleImport=(await importXlsx(zipWrite(Object.fromEntries(archive)))).document.active.tools.conditional;assert.equal(multiRuleImport.length,2);assert.equal(multiRuleImport[0].formula,'=B2<15');assert.equal(multiRuleImport[0].priority,1);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,formula:'=NOT A FORMULA ???'}]},12,8),/formula is invalid/);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,priority:0}]},12,8),/Invalid conditional formatting rule/);assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,stopIfTrue:'yes'}]},12,8),/Invalid conditional formatting rule/);
 assert.equal(ruleColor('B2',12,[{...rule,formula:'=1/0'}],wb),null);
});
test('conditional icon sets calculate ranges, reverse icons, and round-trip native XLSX thresholds',async()=>{
 const wb=new Workbook(12,8);wb.applyEdits([{ref:'B2',raw:'12'},{ref:'B3',raw:'8'},{ref:'B4',raw:'20'},{ref:'C2',raw:'10'},{ref:'C3',raw:'10'},{ref:'C4',raw:'10'}],false);
 const defaults={start:'B2',end:'B4',operator:'iconSet',threshold:0,color:'#000000',iconSet:'3TrafficLights1'};assert.equal(conditionalIcon('B2',12,[defaults],wb)?.index,1);assert.equal(conditionalIcon('B3',8,[defaults],wb)?.index,0);assert.equal(conditionalIcon('B4',20,[defaults],wb)?.index,2);
 assert.equal(conditionalIcon('B3',8,[{...defaults,reverseIcons:true}],wb)?.index,2);assert.equal(conditionalIcon('B4',20,[{...defaults,iconsOnly:true}],wb)?.iconsOnly,true);
 const custom={...defaults,iconSet:'3Arrows',iconThresholds:[{type:'min'},{type:'num',value:10},{type:'num',value:15}]};assert.equal(conditionalIcon('B2',12,[custom],wb)?.index,1);
 const formulaThreshold={...custom,iconThresholds:[{type:'min'},{type:'formula',value:'=C2'},{type:'num',value:15}]};assert.equal(conditionalIcon('B2',12,[formulaThreshold],wb)?.index,1);assert.equal(conditionalIcon('B3',8,[formulaThreshold],wb)?.index,0);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...custom,iconThresholds:[{type:'min'}]}]},12,8),/Invalid conditional icon-set rule/);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...custom,iconThresholds:[{type:'min'},{type:'formula',value:'='},{type:'num',value:15}]}]},12,8),/threshold formula is invalid/);
 const doc=new SpreadsheetDocument(wb);doc.active.tools.conditional.push({...custom,reverseIcons:true,iconsOnly:true,priority:1});doc.active.tools.conditional.push({...formulaThreshold,start:'C2',end:'C4',priority:2});const archive=await zipRead(exportXlsx(doc)),sheetXml=archive.get('xl/worksheets/sheet1.xml');assert.match(sheetXml,/<iconSet iconSet="3Arrows" reverse="1" showValue="0"><cfvo type="min"\/><cfvo type="num" val="10"\/><cfvo type="num" val="15"\/><\/iconSet>/);assert.match(sheetXml,/<cfvo type="formula" val="C2"\/>/);
 const restoredRules=(await importXlsx(exportXlsx(doc))).document.active.tools.conditional,restored=restoredRules[0];assert.equal(restored.operator,'iconSet');assert.equal(restored.iconSet,'3Arrows');assert.deepEqual(restored.iconThresholds,custom.iconThresholds);assert.equal(restored.reverseIcons,true);assert.equal(restored.iconsOnly,true);assert.equal(restored.priority,1);assert.deepEqual(restoredRules[1].iconThresholds,formulaThreshold.iconThresholds);
 const odsRules=(await importOds(exportOds(doc))).document.active.tools.conditional;assert.deepEqual(odsRules[0].iconThresholds,custom.iconThresholds);assert.deepEqual(odsRules[1].iconThresholds,formulaThreshold.iconThresholds);
});
test('custom color-scale thresholds preserve XLSX rules and interpolate across configured stops',async()=>{
 const wb=new Workbook(10,5);wb.applyEdits([{ref:'A1',raw:'0'},{ref:'A2',raw:'25'},{ref:'A3',raw:'50'},{ref:'A4',raw:'75'},{ref:'A5',raw:'100'},{ref:'B1',raw:'0'},{ref:'B2',raw:'50'},{ref:'B3',raw:'100'}],false);
 const rule={start:'A1',end:'A5',operator:'scale',threshold:0,startColor:'#00ff00',midColor:'#ffff00',color:'#ff0000',scaleThresholds:[{type:'num',value:20},{type:'num',value:60},{type:'max'}]};
 assert.equal(ruleColor('A1',0,[rule],wb),'#00ff00');assert.equal(ruleColor('A2',25,[rule],wb),'#20ff00');assert.equal(ruleColor('A3',50,[rule],wb),'#bfff00');assert.equal(ruleColor('A4',75,[rule],wb),'#ff9f00');assert.equal(ruleColor('A5',100,[rule],wb),'#ff0000');
 const twoColor={...rule,start:'B1',end:'B3',midColor:undefined,scaleThresholds:[{type:'num',value:25},{type:'num',value:75}]};assert.equal(ruleColor('B2',50,[twoColor],wb),'#808000');
 const percentile={...rule,start:'A1',end:'A5',scaleThresholds:[{type:'min'},{type:'percentile',value:50},{type:'max'}]},scaleValueCache=new Map();assert.equal(ruleColor('A3',50,[percentile],wb,new Map(),scaleValueCache),'#ffff00');assert.equal(ruleColor('A4',75,[percentile],wb,new Map(),scaleValueCache),'#ff8000');assert.deepEqual(scaleValueCache.get(percentile),[0,25,50,75,100]);
 wb.applyEdits([{ref:'C1',raw:'0'},{ref:'C2',raw:'25'},{ref:'C3',raw:'75'}],false);const formulaStops={...rule,start:'B1',end:'B3',scaleThresholds:[{type:'min'},{type:'formula',value:'=C1'},{type:'max'}]};assert.equal(ruleColor('B2',50,[formulaStops],wb),'#ffaa00');
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,scaleThresholds:[{type:'min'},{type:'max'}]}]},10,5),/thresholds must match/);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,scaleThresholds:[{type:'num',value:Infinity},{type:'num',value:60},{type:'max'}]}]},10,5),/Invalid color-scale threshold value/);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,scaleThresholds:[{type:'min'},{type:'formula',value:'=('},{type:'max'}]}]},10,5),/threshold formula is invalid/);
 const doc=new SpreadsheetDocument(wb);doc.active.tools.conditional.push(rule);const xml=(await zipRead(exportXlsx(doc))).get('xl/worksheets/sheet1.xml');assert.match(xml,/<colorScale><cfvo type="num" val="20"\/><cfvo type="num" val="60"\/><cfvo type="max"\/><color rgb="FF00FF00"\/><color rgb="FFFFFF00"\/><color rgb="FFFF0000"\/><\/colorScale>/);
 const restored=(await importXlsx(exportXlsx(doc))).document.active.tools.conditional[0];assert.deepEqual(restored.scaleThresholds,rule.scaleThresholds);assert.equal(ruleColor('A3',50,[restored],wb),'#bfff00');
 const odsRestored=(await importOds(exportOds(doc))).document.active.tools.conditional[0];assert.deepEqual(odsRestored.scaleThresholds,rule.scaleThresholds);assert.equal(odsRestored.operator,'scale');
});
test('secondary series receive independent value scales',()=>{
 const scales=calculateAxisScales([{name:'Sales',values:[1000,2000]},{name:'Margin',values:[.1,.2]}],[{axis:'primary'},{axis:'secondary'}]);
 assert.deepEqual(scales.primary,{min:0,max:2000});assert.deepEqual(scales.secondary,{min:0,max:.2});
});
test('data bars calculate live value extents and preserve rule semantics in JSON validation',()=>{
 const book=new Workbook(10,5);book.applyEdits([{ref:'A1',raw:'-10'},{ref:'A2',raw:'0'},{ref:'A3',raw:'20'}],false);
 const rule={start:'A1',end:'A3',operator:'dataBar',threshold:0,color:'#2878c8',negativeColor:'#cc4455',barStyle:'solid',showValue:false};
 assert.equal(dataBarFill('A3',20,[rule],book).hideValue,true);
 assert.match(dataBarFill('A1',-10,[rule],book).background,/cc4455/);
 const centered={...rule,barAxis:'middle',barAxisColor:'#123456'};assert.match(dataBarFill('A3',20,[centered],book).background,/calc\(50% - \.5px\).*#123456.*calc\(50% \+ \.5px\)/);assert.match(dataBarFill('A3',20,[rule],book).background,/#808080/);assert.doesNotMatch(dataBarFill('A3',20,[{...centered,barAxis:'none'}],book).background,/#123456/);
 const sized={...rule,start:'B1',end:'B2',barMinLength:20,barMaxLength:80};book.applyEdits([{ref:'B1',raw:'5'},{ref:'B2',raw:'20'}],false);assert.match(dataBarFill('B1',5,[sized],book).background,/35%/);
 assert.equal(validateTools({...freshTools(),conditional:[rule]},10,5).conditional[0].operator,'dataBar');
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,minValue:5,maxValue:2}]},10,5));
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,barMinLength:80,barMaxLength:20}]},10,5));
});
test('data bars evaluate and round-trip typed number, percentile, and formula bounds',async()=>{
 const book=new Workbook(10,5);book.applyEdits([{ref:'A1',raw:'10'},{ref:'A2',raw:'20'},{ref:'A3',raw:'30'},{ref:'A4',raw:'40'}],false);const rule={start:'A1',end:'A4',operator:'dataBar',threshold:0,color:'#2878c8',barLowerThreshold:{type:'percentile',value:25},barUpperThreshold:{type:'formula',value:'=A1+30'}};
 const applied=validateTools({...freshTools(),conditional:[rule]},book.rows,book.cols).conditional[0];assert.deepEqual(applied.barLowerThreshold,{type:'percentile',value:25});assert.deepEqual(applied.barUpperThreshold,{type:'formula',value:'=A1+30'});assert.match(dataBarFill('A2',20,[rule],book).background,/11\.11111111111111%/);
 const untrusted={...rule,barLowerThreshold:{...rule.barLowerThreshold}};const isolated=validateTools({...freshTools(),conditional:[untrusted]},book.rows,book.cols).conditional[0];untrusted.barLowerThreshold.value=90;assert.equal(isolated.barLowerThreshold.value,25);
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,barLowerThreshold:{type:'percent',value:101}}]},book.rows,book.cols),/data-bar lower threshold value/);assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,barUpperThreshold:{type:'formula',value:'=('}}]},book.rows,book.cols),/data-bar upper threshold formula is invalid/);assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,barLowerThreshold:{type:'num',value:35},barUpperThreshold:{type:'num',value:25}}]},book.rows,book.cols),/lower bound must be less than/);assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,minValue:1}]},book.rows,book.cols),/cannot combine legacy fixed limits/);
 const doc=new SpreadsheetDocument(book);doc.active.tools.conditional.push(rule);const xlsx=await importXlsx(exportXlsx(doc)),xlsxRule=xlsx.document.active.tools.conditional[0];assert.deepEqual(xlsxRule.barLowerThreshold,rule.barLowerThreshold);assert.deepEqual(xlsxRule.barUpperThreshold,rule.barUpperThreshold);const xlsxBytes=exportXlsx(doc),xlsxXml=(await zipRead(xlsxBytes)).get('xl/worksheets/sheet1.xml');assert.match(xlsxXml,/<cfvo type="percentile" val="25"\/><cfvo type="formula" val="A1\+30"\/>/);const malformedArchive=await zipRead(xlsxBytes);malformedArchive.set('xl/worksheets/sheet1.xml',xlsxXml.replace('val="A1+30"','val="("'));assert.equal((await importXlsx(zipWrite(Object.fromEntries(malformedArchive)))).document.active.tools.conditional.length,0);
 const odsArchive=await zipRead(exportOds(doc)),content=odsArchive.get('content.xml');assert.match(content,/calcext:type="percentile"/);assert.match(content,/calcext:value="of:=\(\[\.A1\]\+30\)" calcext:type="formula"/);odsArchive.set('content.xml',content.replace(/ gridforge:sheet-objects="[^"]*"/g,''));const odsRule=(await importOds(zipWrite({mimetype:'application/vnd.oasis.opendocument.spreadsheet',...Object.fromEntries(odsArchive)}))).document.active.tools.conditional[0];assert.deepEqual(odsRule.barLowerThreshold,rule.barLowerThreshold);assert.deepEqual(odsRule.barUpperThreshold,{type:'formula',value:'=(A1+30)'});
});
test('conditional numeric samples use sparse populated rows across a full-height worksheet',()=>{
 const book=new Workbook(1_048_576,1);book.applyEdits([{ref:'A1',raw:'10'},{ref:'A2',raw:'=SEQUENCE(3)'},{ref:'A1048576',raw:'9'}],false);assert.deepEqual(book.getNumericValuesInRange('A1','A1048576').sort((a,b)=>a-b),[1,2,3,9,10]);
});
test('malformed XLSX scale and icon thresholds are skipped without rejecting workbook data',async()=>{
 for(const rule of [
  {start:'A1',end:'A3',operator:'scale',threshold:0,startColor:'#ffffff',color:'#0000ff',scaleThresholds:[{type:'min'},{type:'formula',value:'=A1+1'}]},
  {start:'A1',end:'A3',operator:'iconSet',threshold:0,color:'#000000',iconSet:'3Arrows',iconThresholds:[{type:'min'},{type:'percent',value:33},{type:'formula',value:'=A1+1'}]},
 ]){const doc=new SpreadsheetDocument(new Workbook(10,5));doc.active.tools.conditional.push(rule);const archive=await zipRead(exportXlsx(doc)),path='xl/worksheets/sheet1.xml',xml=archive.get(path);archive.set(path,xml.replace('val="A1+1"','val="("'));const imported=await importXlsx(zipWrite(Object.fromEntries(archive)));assert.equal(imported.document.active.tools.conditional.length,0);assert.equal(imported.document.active.book.getRaw('A1'),'');}
});
test('v11.2 tables, data bars, chart axes/trendlines and dimensions roundtrip in XLSX',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));const book=doc.active.book;
 book.applyEdits([{ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Margin'},{ref:'A2',raw:'Jan'},{ref:'B2',raw:'10'},{ref:'C2',raw:'0.1'},{ref:'A3',raw:'Feb'},{ref:'B3',raw:'20'},{ref:'C3',raw:'0.2'}],false);
 doc.active.tools.tables.push({id:'table-v111',name:'Metrics',range:'A1:C3',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false,calculatedColumns:[{column:2,formula:'=[@Revenue]/10',overrides:[]}]});
 doc.active.tools.conditional.push({start:'B2',end:'B3',operator:'dataBar',threshold:0,color:'#2878c8',barStyle:'solid',barAxis:'middle',barAxisColor:'#445566',barMinLength:12,barMaxLength:90});
 doc.active.tools.columnWidths[1]=180;doc.active.tools.rowHeights[2]=42;
 doc.active.tools.charts.push(validateChart({id:'chart-v111',kind:'combo',title:'Revenue and margin',range:'A1:C3',x:0,y:0,width:500,height:300,secondaryYAxisTitle:'Margin %',yAxisMin:0,yAxisMax:2500,secondaryYAxisMin:0,secondaryYAxisMax:1,yAxisNumberFormat:'currency',majorGridlines:true,minorGridlines:true,seriesOptions:[{kind:'column',color:'#214ac7',errorBars:{type:'percentage',value:5}},{axis:'secondary',kind:'line',marker:false,trendline:{type:'linear',forward:1,showR2:true}}]},20,10));
 const archive=await zipRead(exportXlsx(doc)),tableXml=archive.get('xl/tables/table1.xml'),sheetXml=archive.get('xl/worksheets/sheet1.xml'),chartXml=archive.get('xl/charts/chart1.xml');
 assert.match(tableXml,/calculatedColumnFormula/);assert.match(sheetXml,/type="dataBar"/);assert.match(sheetXml,/axisPosition="middle"/);assert.match(sheetXml,/axisColor rgb="FF445566"/);assert.match(sheetXml,/minLength="12"/);assert.match(sheetXml,/maxLength="90"/);assert.match(sheetXml,/customWidth="1"/);assert.match(sheetXml,/customHeight="1"/);assert.match(chartXml,/axId val="404"/);assert.match(chartXml,/trendlineType val="linear"/);
 const xlsxConditional=(await importXlsx(exportXlsx(doc))).document.active.tools.conditional[0];assert.equal(xlsxConditional.barAxis,'middle');assert.equal(xlsxConditional.barAxisColor,'#445566');assert.equal(xlsxConditional.barMinLength,12);assert.equal(xlsxConditional.barMaxLength,90);
 assert.match(chartXml,/<c:min val="0"\/>/);assert.match(chartXml,/<c:max val="2500"\/>/);assert.match(chartXml,/<c:minorGridlines\/>/);assert.match(chartXml,/<c:numFmt formatCode="\$#,##0\.00"/);assert.match(chartXml,/srgbClr val="214AC7"/);assert.match(chartXml,/symbol val="none"/);assert.match(chartXml,/errValType val="percentage"/);
 const restored=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(restored.active.tools.columnWidths[1],180);assert.equal(restored.active.tools.rowHeights[2],42);assert.equal(restored.active.tools.tables[0].calculatedColumns[0].formula,'=[@Revenue]/10');
});

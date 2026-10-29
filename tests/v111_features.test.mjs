import test from 'node:test';
import assert from 'node:assert/strict';
import {fitTrendline} from '../dist/engine/trendlines.js';
import {Workbook} from '../dist/engine/workbook.js';
import {dataBarFill,validateTools,freshTools} from '../dist/engine/sheettools.js';
import {validateChart,chartSvg,calculateAxisScales} from '../dist/engine/charts.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';

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
 const wb=new Workbook(20,10);wb.applyEdits([{ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Margin'},{ref:'A2',raw:'Jan'},{ref:'B2',raw:'100'},{ref:'C2',raw:'20'},{ref:'A3',raw:'Feb'},{ref:'B3',raw:'200'},{ref:'C3',raw:'30'}],false);const svg=chartSvg(wb,chart);assert.match(svg,/stroke="#d12b52"/);assert.doesNotMatch(svg,/fill="#d12b52"\/>/);assert.match(svg,/\$100/);assert.match(svg,/stroke="#f1f4f8"/);
 assert.throws(()=>validateChart({...chart,yAxisMin:501},20,10),/minimum must be less/);
});
test('series error bars render, validate and persist through XLSX chart parts',async()=>{
 const chart=validateChart({id:'chart-errors',kind:'line',title:'Uncertainty',range:'A1:B4',x:0,y:0,width:500,height:300,seriesOptions:[{errorBars:{type:'fixed',value:2}}]},12,8);
 const wb=new Workbook(12,8);wb.applyEdits([{ref:'A1',raw:'Period'},{ref:'B1',raw:'Value'},{ref:'A2',raw:'1'},{ref:'B2',raw:'10'},{ref:'A3',raw:'2'},{ref:'B3',raw:'14'},{ref:'A4',raw:'3'},{ref:'B4',raw:'12'}],false);
 assert.match(chartSvg(wb,chart),/<g stroke="#4472c4"[^>]*><line/);
 assert.throws(()=>validateChart({...chart,seriesOptions:[{errorBars:{type:'fixed',value:-1}}]},12,8),/error bars/);
 const doc=new SpreadsheetDocument(wb);doc.active.tools.charts.push(chart);const zip=await zipRead(exportXlsx(doc)),xml=zip.get('xl/charts/chart1.xml');assert.match(xml,/<c:errValType val="fixedVal"\/><c:val val="2"\/>/);
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
 assert.equal(validateTools({...freshTools(),conditional:[rule]},10,5).conditional[0].operator,'dataBar');
 assert.throws(()=>validateTools({...freshTools(),conditional:[{...rule,minValue:5,maxValue:2}]},10,5));
});
test('v11.2 tables, data bars, chart axes/trendlines and dimensions roundtrip in XLSX',async()=>{
 const doc=new SpreadsheetDocument(new Workbook(20,10));const book=doc.active.book;
 book.applyEdits([{ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Margin'},{ref:'A2',raw:'Jan'},{ref:'B2',raw:'10'},{ref:'C2',raw:'0.1'},{ref:'A3',raw:'Feb'},{ref:'B3',raw:'20'},{ref:'C3',raw:'0.2'}],false);
 doc.active.tools.tables.push({id:'table-v111',name:'Metrics',range:'A1:C3',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false,calculatedColumns:[{column:2,formula:'=[@Revenue]/10',overrides:[]}]});
 doc.active.tools.conditional.push({start:'B2',end:'B3',operator:'dataBar',threshold:0,color:'#2878c8',barStyle:'solid'});
 doc.active.tools.columnWidths[1]=180;doc.active.tools.rowHeights[2]=42;
 doc.active.tools.charts.push(validateChart({id:'chart-v111',kind:'combo',title:'Revenue and margin',range:'A1:C3',x:0,y:0,width:500,height:300,secondaryYAxisTitle:'Margin %',yAxisMin:0,yAxisMax:2500,secondaryYAxisMin:0,secondaryYAxisMax:1,yAxisNumberFormat:'currency',majorGridlines:true,minorGridlines:true,seriesOptions:[{kind:'column',color:'#214ac7',errorBars:{type:'percentage',value:5}},{axis:'secondary',kind:'line',marker:false,trendline:{type:'linear',forward:1,showR2:true}}]},20,10));
 const archive=await zipRead(exportXlsx(doc)),tableXml=archive.get('xl/tables/table1.xml'),sheetXml=archive.get('xl/worksheets/sheet1.xml'),chartXml=archive.get('xl/charts/chart1.xml');
 assert.match(tableXml,/calculatedColumnFormula/);assert.match(sheetXml,/type="dataBar"/);assert.match(sheetXml,/customWidth="1"/);assert.match(sheetXml,/customHeight="1"/);assert.match(chartXml,/axId val="404"/);assert.match(chartXml,/trendlineType val="linear"/);
 assert.match(chartXml,/<c:min val="0"\/>/);assert.match(chartXml,/<c:max val="2500"\/>/);assert.match(chartXml,/<c:minorGridlines\/>/);assert.match(chartXml,/<c:numFmt formatCode="\$#,##0\.00"/);assert.match(chartXml,/srgbClr val="214AC7"/);assert.match(chartXml,/symbol val="none"/);assert.match(chartXml,/errValType val="percentage"/);
 const restored=SpreadsheetDocument.fromJSON(doc.toJSON());assert.equal(restored.active.tools.columnWidths[1],180);assert.equal(restored.active.tools.rowHeights[2],42);assert.equal(restored.active.tools.tables[0].calculatedColumns[0].formula,'=[@Revenue]/10');
});

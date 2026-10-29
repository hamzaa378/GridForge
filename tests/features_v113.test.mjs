import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {chartSvg,treemapData,stockData} from '../dist/engine/charts.js';
import {buildDescriptiveStatistics} from '../dist/engine/descriptive_stats.js';

test('funnel chart renders proportional sequential stages and rejects negative values gracefully',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Stage'},{ref:'B1',raw:'Users'},{ref:'A2',raw:'Visitors'},{ref:'B2',raw:'100'},{ref:'A3',raw:'Trials'},{ref:'B3',raw:'45'},{ref:'A4',raw:'Paid'},{ref:'B4',raw:'12'}]);
  const chart={id:'chart-funnel',kind:'funnel',title:'Conversion',range:'A1:B4',x:0,y:0,width:500,height:310};const svg=chartSvg(book,chart);assert.equal((svg.match(/<path /g)??[]).length,3);assert.match(svg,/Visitors/);
  book.setCell('B3','-2');assert.match(chartSvg(book,chart),/Funnel charts require non-negative values/);
});

test('treemap groups hierarchy paths, adds duplicate categories, and renders weighted tiles',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Product'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'A'},{ref:'C2',raw:'10'},{ref:'A3',raw:'West'},{ref:'B3',raw:'A'},{ref:'C3',raw:'5'},{ref:'A4',raw:'West'},{ref:'B4',raw:'B'},{ref:'C4',raw:'4'},{ref:'A5',raw:'East'},{ref:'B5',raw:'A'},{ref:'C5',raw:'6'}]);
  const data=treemapData(book,'A1:C5');assert.deepEqual(data.labels,['West › A','East › A','West › B']);assert.deepEqual(data.series[0].values,[15,6,4]);
  const svg=chartSvg(book,{id:'chart-tree',kind:'treemap',title:'Sales by product',range:'A1:C5',x:0,y:0,width:500,height:310,dataLabels:true});assert.equal((svg.match(/<rect /g)??[]).length,4);assert.match(svg,/West › A/);
  book.setCell('C5','-1');assert.throws(()=>treemapData(book,'A1:C5'),/non-negative/);
});

test('Descriptive Statistics writes sample metrics to a separate report range',()=>{
  const book=new Workbook(30,10);book.applyEdits([{ref:'A1',raw:'1'},{ref:'A2',raw:'2'},{ref:'A3',raw:'2'},{ref:'A4',raw:'3'},{ref:'A5',raw:'4'},{ref:'A6',raw:'text'},{ref:'A7',raw:'=1/0'}]);
  const report=buildDescriptiveStatistics(book,'A1:A7','D1');assert.equal(report.count,5);assert.equal(report.outputRange,'D1:E14');book.applyEdits(report.edits);
  assert.equal(book.getValue('D1').value,'Statistic');assert.equal(book.getValue('E2').value,2.4);assert.equal(book.getValue('E4').value,2);assert.equal(book.getValue('E5').value,2);assert.equal(book.getValue('E7').value,1.3);assert.equal(book.getValue('E14').value,5);assert.equal(book.getRaw('A1'),'1');
  assert.throws(()=>buildDescriptiveStatistics(book,'A1:A7','A4'),/outside the selected source range/);assert.throws(()=>buildDescriptiveStatistics(book,'A6:A7','D1'),/no numeric values/);
});

test('stock charts map OHLC columns by header name and show rising/falling candles',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Date'},{ref:'B1',raw:'Close'},{ref:'C1',raw:'Open'},{ref:'D1',raw:'High'},{ref:'E1',raw:'Low'},{ref:'A2',raw:'Mon'},{ref:'B2',raw:'12'},{ref:'C2',raw:'10'},{ref:'D2',raw:'14'},{ref:'E2',raw:'9'},{ref:'A3',raw:'Tue'},{ref:'B3',raw:'9'},{ref:'C3',raw:'12'},{ref:'D3',raw:'13'},{ref:'E3',raw:'8'}]);
  const data=stockData(book,'A1:E3');assert.deepEqual(data.labels,['Mon','Tue']);assert.deepEqual(data.series.map(series=>series.values),[[10,12],[14,13],[9,8],[12,9]]);
  const svg=chartSvg(book,{id:'chart-stock',kind:'stock',title:'Prices',range:'A1:E3',x:0,y:0,width:500,height:300});assert.match(svg,/#18865b/);assert.match(svg,/#d44949/);
  book.setCell('D2','11');assert.match(chartSvg(book,{id:'chart-stock',kind:'stock',title:'Prices',range:'A1:E3',x:0,y:0,width:500,height:300}),/Invalid stock values/);
});

test('sunburst chart aggregates each hierarchy level and renders nested ring segments',()=>{
  const book=new Workbook(20,10);book.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Product'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'A'},{ref:'C2',raw:'10'},{ref:'A3',raw:'West'},{ref:'B3',raw:'B'},{ref:'C3',raw:'5'},{ref:'A4',raw:'East'},{ref:'B4',raw:'A'},{ref:'C4',raw:'8'}]);
  const svg=chartSvg(book,{id:'chart-sunburst',kind:'sunburst',title:'Sales mix',range:'A1:C4',x:0,y:0,width:520,height:330});assert.equal((svg.match(/<path /g)??[]).length,5);assert.match(svg,/West/);assert.match(svg,/East/);assert.match(svg,/<circle/);
  book.setCell('C3','-1');assert.match(chartSvg(book,{id:'chart-sunburst',kind:'sunburst',title:'Sales mix',range:'A1:C4',x:0,y:0,width:520,height:330}),/Hierarchy chart values must be non-negative/);
  const single=new Workbook(10,5);single.applyEdits([{ref:'A1',raw:'Region'},{ref:'B1',raw:'Product'},{ref:'C1',raw:'Sales'},{ref:'A2',raw:'West'},{ref:'B2',raw:'A'},{ref:'C2',raw:'5'}]);const fullRing=chartSvg(single,{id:'chart-sunburst-single',kind:'sunburst',title:'One group',range:'A1:C2',x:0,y:0,width:500,height:300});assert.match(fullRing,/fill-rule="evenodd"/);assert.doesNotMatch(fullRing,/<path d="<path/);
});

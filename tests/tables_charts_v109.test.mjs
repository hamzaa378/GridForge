import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {TABLE_STYLES,validateTable,tableAppearance,tableRange,tableColumns} from '../dist/engine/tables.js';
import {validateTools} from '../dist/engine/sheettools.js';
import {CHART_KINDS,CHART_PALETTES,chartSvg,validateChart} from '../dist/engine/charts.js';
import {exportXlsx} from '../dist/engine/xlsx.js';
import {zipRead} from '../dist/engine/zip.js';
import {cleanStyle} from '../dist/engine/format.js';
const w=new Workbook(400,52);
w.applyEdits([{ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Costs'},...['Jan','Feb','Mar','Apr'].flatMap((s,i)=>[{ref:`A${i+2}`,raw:s},{ref:`B${i+2}`,raw:String(100*(i+1))},{ref:`C${i+2}`,raw:String(70*(i+1))}])],false);
const def={id:'table-sales',name:'Sales2026',range:'A1:C5',style:'TableStyleMedium2',header:true,bandedRows:true,bandedColumns:false,firstColumn:true,lastColumn:false};
function fresh(){const d=new SpreadsheetDocument(Workbook.fromJSON(w.toJSON()));d.active.tools.tables=[{...def}];return d;}
test('v10.9 ten Excel-compatible table style names',()=>{assert.equal(TABLE_STYLES.length,10);assert.ok(TABLE_STYLES.includes('TableStyleDark4'));});
test('table validation rejects bad rectangles, overlapping tables and duplicate names',()=>{
 assert.equal(validateTable(def,400,52).name,'Sales2026');
 assert.throws(()=>validateTable({...def,range:'C5:A1'},400,52));
 assert.throws(()=>validateTable({...def,name:'A1'},400,52));
 assert.throws(()=>validateTools({...fresh().active.tools,tables:[def,{...def,id:'table-two',name:'NotSame',range:'B3:E9'}]},400,52),/overlapping|duplicate/i);
});
test('table formatting honors header, bands, and emphasized first column',()=>{
 const b=tableRange(def.range,400,52);
 assert.equal(tableAppearance(def,b,0,0).bold,true);
 assert.notEqual(tableAppearance(def,b,1,1).background,tableAppearance(def,b,2,1).background);
 assert.equal(tableAppearance(def,b,2,0).bold,true);
 assert.equal(tableAppearance(def,b,8,0),null);
});
test('table headers derive from real spreadsheet values without destroying data',()=>{
 assert.deepEqual(tableColumns(w,def),['Month','Revenue','Costs']);
 assert.equal(w.getRaw('B2'),'100');
});
test('tables and chart definitions persist through GridForge JSON including sheet duplication',()=>{
 const d=fresh();d.active.tools.charts=[{id:'chart-a',kind:'combo',theme:'forest',title:'Revenue and costs',range:'A1:C5',x:200,y:120,width:500,height:310}];
 const clone=SpreadsheetDocument.fromJSON(d.toJSON());
 assert.equal(clone.active.tools.tables[0].name,'Sales2026');
 assert.equal(clone.active.tools.charts[0].theme,'forest');
 const copy=clone.duplicateSheet();assert.equal(copy.tools.tables.length,1);
});
test('new chart varieties and chart theme selectors validate and render offline',()=>{
 for(const kind of ['percent-column','percent-bar','combo','radar']){
   const chart=validateChart({id:`chart-${kind}`,kind,theme:'forest',title:kind,range:'A1:C5',x:0,y:0,width:500,height:330},400,52);
   const svg=chartSvg(w,chart);assert.match(svg,/^<svg/);assert.match(svg,/Revenue|Costs|Jan/);
 }
 assert.throws(()=>validateChart({id:'chart-bad',kind:'column',theme:'unknown',title:'?',range:'A1:C5',x:0,y:0,width:500,height:300},400,52));
 assert.equal(CHART_KINDS.length,21);assert.equal(Object.keys(CHART_PALETTES).length,6);
});
test('expanded font families and sizes remain validated and safe',()=>{
 assert.equal(cleanStyle({fontFamily:'Cambria',fontSize:36,color:'#aa33cc'}).fontFamily,'Cambria');
 assert.throws(()=>cleanStyle({fontFamily:'<script>'}));
 assert.throws(()=>cleanStyle({fontSize:99}));
});
test('XLSX emits native named table, relationship, style, and correct worksheet values',async()=>{
 const d=fresh(),archive=await zipRead(exportXlsx(d));
 const main=archive.get('xl/tables/table1.xml');assert.ok(main);
 assert.match(main,/name="Sales2026"/);assert.match(main,/TableStyleMedium2/);assert.match(main,/name="Revenue"/);
 assert.match(archive.get('xl/worksheets/sheet1.xml'),/<tableParts/);
 assert.match(archive.get('xl/worksheets/_rels/sheet1.xml.rels'),/\/tables\/table1.xml/);
 assert.match(archive.get('[Content_Types].xml'),/spreadsheetml.table\+xml/);
 assert.match(archive.get('xl/worksheets/sheet1.xml'),/<v>100<\/v>/);
});
test('charts and tables share correct per-sheet relationship ids in XLSX',async()=>{
 const d=fresh();
 d.active.tools.charts=['percent-column','percent-bar','combo','radar'].map((kind,i)=>({id:`chart-${i}`,kind,theme:'sunset',title:kind,range:'A1:C5',x:120,y:i*320,width:480,height:310}));
 const archive=await zipRead(exportXlsx(d));
 const relationships=archive.get('xl/worksheets/_rels/sheet1.xml.rels');
 assert.match(relationships,/Id="rId1"[^>]+\/drawing"/);
 assert.match(relationships,/Id="rId2"[^>]+\/table"/);
 assert.match(archive.get('xl/charts/chart1.xml'),/percentStacked/);
 assert.match(archive.get('xl/charts/chart2.xml'),/percentStacked/);
 assert.match(archive.get('xl/charts/chart3.xml'),/<c:lineChart>/);
 assert.match(archive.get('xl/charts/chart4.xml'),/<c:radarChart>/);
 assert.equal([...archive.keys()].filter(n=>/^xl\/charts\/chart\d+\.xml$/.test(n)).length,4);
});

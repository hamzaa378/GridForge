import {writeFile} from 'node:fs/promises';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {exportXlsx} from '../dist/engine/xlsx.js';

const book=new Workbook(400,52);
book.applyEdits([
 {ref:'A1',raw:'Month'},{ref:'B1',raw:'Revenue'},{ref:'C1',raw:'Margin'},{ref:'D1',raw:'Profit'},
 ...[['Jan',12000,.12],['Feb',14500,.15],['Mar',13800,.13],['Apr',17200,.18],['May',19600,.2]].flatMap(([month,revenue,margin],i)=>[
  {ref:`A${i+2}`,raw:month},{ref:`B${i+2}`,raw:String(revenue)},{ref:`C${i+2}`,raw:String(margin)},{ref:`D${i+2}`,raw:`=[@Revenue]*[@Margin]`}
 ])
],false);
const doc=new SpreadsheetDocument(book),sheet=doc.active;
sheet.tools.tables.push({id:'table-monthly',name:'MonthlyPerformance',range:'A1:D6',style:'TableStyleMedium4',header:true,bandedRows:true,bandedColumns:false,firstColumn:false,lastColumn:false,calculatedColumns:[{column:3,formula:'=[@Revenue]*[@Margin]',overrides:[]}]});
book.setTableContext(sheet.tools.tables);
sheet.tools.conditional.push({start:'B2',end:'B6',operator:'dataBar',threshold:0,color:'#2f80c9',negativeColor:'#c84b54',barStyle:'gradient',showValue:true});
sheet.tools.columnWidths={0:100,1:150,2:115,3:150};sheet.tools.rowHeights={0:30};
sheet.tools.charts.push({id:'chart-performance',kind:'combo',title:'Revenue and profit margin',range:'A1:C6',x:220,y:80,width:560,height:340,legend:'bottom',xAxisTitle:'Month',yAxisTitle:'Revenue ($)',secondaryYAxisTitle:'Profit margin',seriesOptions:[
 {axis:'primary',kind:'column'},
 {axis:'secondary',kind:'line',trendline:{type:'linear',forward:1,showEquation:true,showR2:true}}
]});
const base=new URL('../examples/',import.meta.url);
await writeFile(new URL('GridForge_v11_1_Advanced_Charts_Tables_Demo.json',base),JSON.stringify(doc.toJSON(),null,2));
await writeFile(new URL('GridForge_v11_1_Advanced_Charts_Tables_Demo.xlsx',base),exportXlsx(doc));
console.log('Built v11.1 JSON and XLSX examples.');

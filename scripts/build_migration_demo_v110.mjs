/** Reproducible GridForge-generated XLSX specimen (uses production export code). */
import {writeFileSync} from 'node:fs';
import {Workbook} from '../dist/engine/workbook.js';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {expandTablesAfterEdit} from '../dist/engine/table_growth.js';
import {exportXlsx} from '../dist/engine/xlsx.js';

const book=new Workbook(400,52),doc=new SpreadsheetDocument(book);
const samples=[['Jan',100,4],['Feb',120,3],['Mar',95,5],['Apr',150,6],['May',175,2]];
book.applyEdits([
 ...['Month','Revenue','Units','Total'].map((raw,i)=>({ref:`${'ABCD'[i]}1`,raw})),
 ...samples.flatMap((row,i)=>row.map((val,j)=>({ref:`${'ABC'[j]}${i+2}`,raw:String(val)})))
]);
doc.active.tools.tables=[{id:'table-monthly',name:'SalesMetrics',range:'A1:D6',style:'TableStyleMedium4',header:true,bandedRows:true,bandedColumns:false,firstColumn:true,lastColumn:false}];
book.setTableContext(doc.active.tools.tables);
for(let r=2;r<=6;r++)book.setCell(`D${r}`,'=[@Revenue]*[@Units]');
book.setCell('F1','Total sales');book.setCell('F2','=SUM(SalesMetrics[Total])');
doc.active.tools.charts=[
 {id:'chart-monthly-line',kind:'line',theme:'forest',title:'Monthly revenue and totals',range:'A1:D6',x:550,y:90,width:530,height:300,legend:'bottom',xAxisTitle:'Month',yAxisTitle:'Sales',dataLabels:false},
 {id:'chart-monthly-bars',kind:'column',theme:'office',title:'Sales by month',range:'A1:D6',x:550,y:450,width:530,height:290,legend:'bottom',dataLabels:false}
];
const edits=[{ref:'A7',raw:'Jun'},{ref:'B7',raw:'200'},{ref:'C7',raw:'3'},{ref:'D7',raw:'=[@Revenue]*[@Units]'}];
book.applyEdits(edits);
const grown=expandTablesAfterEdit(book,doc.active.tools,edits);
if(grown.length)book.setTableContext(doc.active.tools.tables);
doc.active.tools.filter={startRow:0,endRow:6,startCol:0,endCol:3,column:0,query:'',criteria:[
 {column:0,values:['Mar','Jun']},{column:2,values:['3','5']}
]};
doc.active.tools.conditional=[{start:'B2',end:'B7',operator:'scale',threshold:0,startColor:'#f4fbf6',color:'#2d936b'}];
if(book.getValue('F2').value!==3085)throw new Error('Demo total unexpectedly incorrect');
if(book.getValue('D7').value!==600)throw new Error('Current-row structured formula failed');
const filename='examples/GridForge_v11_0_Table_Migration_Demo';
writeFileSync(filename+'.json',JSON.stringify(doc.toJSON(),null,2)+'\n');
writeFileSync(filename+'.xlsx',exportXlsx(doc));
console.log(`PASS: XLSX and JSON demo, ${grown.length} expanded table(s), ${doc.active.tools.charts.length} native charts, 2 column filters, live color scale; expected total ${book.getValue('F2').value}`);

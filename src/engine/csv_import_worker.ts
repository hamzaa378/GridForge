import {SpreadsheetDocument} from './document.js';
import {Workbook,type Edit} from './workbook.js';
import {colToName} from './parser.js';
import {parseCSV} from './csv.js';
import {MAX_WORKSHEET_COLUMNS,MAX_WORKSHEET_ROWS,MAX_WORKSHEET_POPULATED_CELLS} from './limits.js';

type Request={requestId:string;bytes:ArrayBuffer};
type Response={requestId:string;kind:'progress';message:string}|{requestId:string;kind:'result';document:unknown;clipped?:{rows:number;columns:number}}|{requestId:string;kind:'error';message:string};
const scope=self as unknown as {onmessage:((event:MessageEvent<Request>)=>void)|null;postMessage:(message:Response)=>void};
const MAX_INPUT_BYTES=64*1024*1024,MAX_SOURCE_CELLS=5_000_000,MAX_NONEMPTY_CELLS=MAX_WORKSHEET_POPULATED_CELLS;

scope.onmessage=event=>{
  const {requestId,bytes}=event.data;
  try{
    if(!(bytes instanceof ArrayBuffer))throw Error('CSV import received invalid file data.');
    if(bytes.byteLength>MAX_INPUT_BYTES)throw Error('CSV files are limited to 64 MB. Split the file into smaller parts and import them separately.');
    const source=new TextDecoder('utf-8',{fatal:false}).decode(bytes),rows=parseCSV(source);
    let longest=0,nonempty=0;
    for(const row of rows){longest=Math.max(longest,row.length);for(const value of row)if(value!=='')nonempty++;}
    const sourceCells=rows.reduce((total,row)=>total+row.length,0);
    if(sourceCells>MAX_SOURCE_CELLS)throw Error('This CSV has more than 5,000,000 fields. Split it into smaller files before importing.');
    if(nonempty>MAX_NONEMPTY_CELLS)throw Error(`This CSV has more than ${MAX_NONEMPTY_CELLS.toLocaleString()} nonempty cells. Split it into smaller files before importing.`);
    const rowCount=Math.max(400,Math.min(MAX_WORKSHEET_ROWS,rows.length)),columnCount=Math.max(52,Math.min(MAX_WORKSHEET_COLUMNS,longest)),workbook=new Workbook(rowCount,columnCount),batch:Edit[]=[],formulas:string[]=[];
    for(let row=0;row<Math.min(rows.length,workbook.rows);row++){
      const values=rows[row]!;
      for(let column=0;column<Math.min(values.length,workbook.cols);column++){
        const raw=values[column]!;
        if(raw!==''){
          const ref=`${colToName(column)}${row+1}`;
          batch.push({ref,raw});
          if(raw.startsWith('='))formulas.push(ref);
          if(batch.length>=4096){workbook.applyEdits(batch,false,false);batch.length=0;}
        }
      }
      if((row+1)%2000===0)scope.postMessage({requestId,kind:'progress',message:`Preparing CSV rows: ${row+1} of ${rows.length}…`});
    }
    if(batch.length)workbook.applyEdits(batch,false,false);
    for(let index=0;index<formulas.length;index++){
      workbook.getValue(formulas[index]!);
      if((index+1)%5000===0)scope.postMessage({requestId,kind:'progress',message:`Calculating CSV formulas: ${index+1} of ${formulas.length}…`});
    }
    const document=new SpreadsheetDocument(workbook);
    scope.postMessage({requestId,kind:'result',document:document.toJSON(),...(rows.length>workbook.rows||longest>workbook.cols?{clipped:{rows:rows.length,columns:longest}}:{})});
  }catch(error){scope.postMessage({requestId,kind:'error',message:error instanceof Error?error.message:String(error)});}
};

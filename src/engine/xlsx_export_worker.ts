import {SpreadsheetDocument} from './document.js';
import {exportXlsx} from './xlsx.js';

type Request={requestId:string;document:unknown};
type Response={requestId:string;kind:'result';bytes:ArrayBuffer}|{requestId:string;kind:'error';message:string};

const workerScope=self as unknown as {
  onmessage:((event:MessageEvent<Request>)=>void)|null;
  postMessage:(message:Response,transfer?:Transferable[])=>void;
};

workerScope.onmessage=event=>{
  const {requestId,document}=event.data;
  try{
    const workbook=SpreadsheetDocument.fromJSON(document);
    const bytes=exportXlsx(workbook);
    const buffer=bytes.slice().buffer;
    workerScope.postMessage({requestId,kind:'result',bytes:buffer},[buffer]);
  }catch(error){
    workerScope.postMessage({requestId,kind:'error',message:error instanceof Error?error.message:String(error)});
  }
};

import {configureXlsxXmlParser} from './xlsx.js';
import {importOds} from './ods.js';

type Request={requestId:string;bytes:ArrayBuffer};
type Response={requestId:string;kind:'progress';message:string}|{requestId:string;kind:'result';document:unknown;warnings:string[]}|{requestId:string;kind:'error';message:string};
const scope=self as unknown as {onmessage:((event:MessageEvent<Request>)=>void)|null;postMessage:(message:Response)=>void};

scope.onmessage=async event=>{
 const {requestId,bytes}=event.data;
 try{
  type ParserConstructor=new(options?:{onError?:(level:string,message:string)=>void})=>{parseFromString:(value:string,mime:string)=>Document};
  const moduleUrl=new URL('../vendor/xmldom.js',import.meta.url).href;
  const loaded=await import(moduleUrl) as {DOMParser?:ParserConstructor;default?:{DOMParser?:ParserConstructor}};
  const XmlParser=loaded.DOMParser??loaded.default?.DOMParser;
  if(!XmlParser)throw Error('The background XML parser did not load.');
  configureXlsxXmlParser(source=>{let failure='';const parser=new XmlParser({onError:(level,message)=>{if(level==='error'||level==='fatalError')failure=message;}}),xml=parser.parseFromString(source,'application/xml');if(failure)throw Error(`Invalid XML inside ODS: ${failure}`);return xml;});
  const result=await importOds(new Uint8Array(bytes),message=>scope.postMessage({requestId,kind:'progress',message}));
  scope.postMessage({requestId,kind:'result',document:result.document.toJSON(),warnings:result.warnings});
 }catch(error){scope.postMessage({requestId,kind:'error',message:error instanceof Error?error.message:String(error)});}
};

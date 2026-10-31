import {configureXlsxXmlParser,importXlsx} from './xlsx.js';

type Request={requestId:string;bytes:ArrayBuffer};
type Response=
  |{requestId:string;kind:'progress';message:string}
  |{requestId:string;kind:'result';document:unknown;warnings:string[]}
  |{requestId:string;kind:'error';message:string};

const workerScope=self as unknown as {
  onmessage:((event:MessageEvent<Request>)=>void)|null;
  postMessage:(message:Response)=>void;
};

workerScope.onmessage=async event=>{
  const {requestId,bytes}=event.data;
  try{
    const moduleUrl=new URL('../vendor/xmldom.js',import.meta.url).href;
    const xmlModule=await import(moduleUrl) as {DOMParser?:new(options?:{onError?:(level:string,message:string)=>void})=>{parseFromString:(text:string,mimeType:string)=>Document};default?:{DOMParser?:new(options?:{onError?:(level:string,message:string)=>void})=>{parseFromString:(text:string,mimeType:string)=>Document}}};
    const XmlDomParser=xmlModule.DOMParser??xmlModule.default?.DOMParser;
    if(!XmlDomParser)throw Error('The background XML parser did not load.');
    configureXlsxXmlParser(text=>{let parseError='';const parser=new XmlDomParser({onError:(level,message)=>{if(level==='error'||level==='fatalError')parseError=message;}}),document=parser.parseFromString(text,'application/xml');if(parseError)throw Error(`Invalid XML inside XLSX: ${parseError}`);return document;});
    const imported=await importXlsx(new Uint8Array(bytes),{onProgress:message=>workerScope.postMessage({requestId,kind:'progress',message})});
    workerScope.postMessage({requestId,kind:'result',document:imported.document.toJSON(),warnings:imported.warnings});
  }catch(error){
    workerScope.postMessage({requestId,kind:'error',message:error instanceof Error?error.message:String(error)});
  }
};

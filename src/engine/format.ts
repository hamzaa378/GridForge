import type {CellResult} from './workbook.js';
import {display} from './workbook.js';
export type NumberFormat='general'|'number'|'integer'|'currency'|'percent';
export type CellAlign='left'|'center'|'right';
export interface CellStyle { bold?:boolean; italic?:boolean; underline?:boolean; strike?:boolean; wrap?:boolean; fontFamily?:string; fontSize?:number; border?:'all'|'bottom'|'top'|'left'|'right'; color?:string; background?:string; align?:CellAlign; verticalAlign?:'top'|'center'; indent?:number; numberFormat?:NumberFormat; dateFormatCode?:string; customNumberFormatCode?:string }
export type StyleEdit={ref:string;style:CellStyle};
const validNumbers=new Set(['general','number','integer','currency','percent']);
const validAlign=new Set(['left','center','right']);
export function cleanStyle(value:unknown):CellStyle {
  if(!value||typeof value!=='object'||Array.isArray(value))throw new Error('Cell formatting must be an object.');
  const source=value as Record<string,unknown>, out:CellStyle={};
  for(const key of ['bold','italic','underline','strike','wrap'] as const)if(source[key]!==undefined){if(typeof source[key]!=='boolean')throw new Error(`Invalid ${key} formatting.`);if(source[key])out[key]=true;}
  for(const key of ['color','background'] as const)if(source[key]!==undefined){if(typeof source[key]!=='string'||!/^#[0-9a-fA-F]{6}$/.test(source[key]))throw new Error(`Invalid ${key}: use a six-digit hex color.`);out[key]=source[key];}
  if(source.align!==undefined){if(typeof source.align!=='string'||!validAlign.has(source.align))throw new Error('Invalid alignment.');if(source.align!=='left')out.align=source.align as CellAlign;}
  if(source.verticalAlign!==undefined){if(source.verticalAlign!=='top'&&source.verticalAlign!=='center')throw new Error('Invalid vertical alignment.');out.verticalAlign=source.verticalAlign;}
  if(source.indent!==undefined){if(typeof source.indent!=='number'||!Number.isInteger(source.indent)||source.indent<0||source.indent>250)throw new Error('Cell indent must be between 0 and 250.');if(source.indent>0)out.indent=source.indent;}
  if(source.numberFormat!==undefined){if(typeof source.numberFormat!=='string'||!validNumbers.has(source.numberFormat))throw new Error('Invalid numeric format.');if(source.numberFormat!=='general')out.numberFormat=source.numberFormat as NumberFormat;}
  if(source.dateFormatCode!==undefined){if(typeof source.dateFormatCode!=='string'||source.dateFormatCode.length>255||/[\u0000-\u001f]/.test(source.dateFormatCode)||!isExcelDateNumberFormat(source.dateFormatCode))throw new Error('Invalid Excel date number format.');out.dateFormatCode=source.dateFormatCode;}
  if(source.customNumberFormatCode!==undefined){if(typeof source.customNumberFormatCode!=='string'||!source.customNumberFormatCode||source.customNumberFormatCode.length>255||/[\u0000-\u001f]/.test(source.customNumberFormatCode))throw new Error('Custom number format must contain 1–255 printable characters.');out.customNumberFormatCode=source.customNumberFormatCode;}
  if(source.fontFamily!==undefined){if(typeof source.fontFamily!=='string'||!['Aptos','Calibri','Arial','Segoe UI','Georgia','Consolas','Verdana','Tahoma','Trebuchet MS','Times New Roman','Cambria','Garamond'].includes(source.fontFamily))throw Error('Invalid font family.');out.fontFamily=source.fontFamily;}
  if(source.fontSize!==undefined){if(typeof source.fontSize!=='number'||!Number.isInteger(source.fontSize)||source.fontSize<8||source.fontSize>36)throw Error('Font size must be between 8 and 36.');out.fontSize=source.fontSize;}
  if(source.border!==undefined){if(!['all','bottom','top','left','right'].includes(String(source.border)))throw Error('Invalid border style.');out.border=source.border as CellStyle['border'];}
  return out;
}
export function styledDisplay(value:CellResult,style:CellStyle):string {
  if(value.error)return display(value);
  if(style.customNumberFormatCode&&typeof value.value==='string'){const formatted=formatExcelText(value.value,style.customNumberFormatCode);if(formatted!==null)return formatted;}
  if(typeof value.value!=='number'||!Number.isFinite(value.value))return display(value);
  const n=value.value;
  if(style.dateFormatCode){const formatted=formatExcelDateSerial(n,style.dateFormatCode);if(formatted!==null)return formatted;}
  if(style.customNumberFormatCode){const formatted=formatExcelNumber(n,style.customNumberFormatCode);if(formatted!==null)return formatted;}
  switch(style.numberFormat){
    case 'integer':return new Intl.NumberFormat('en-US',{maximumFractionDigits:0}).format(n);
    case 'number':return new Intl.NumberFormat('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}).format(n);
    case 'currency':return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
    case 'percent':return new Intl.NumberFormat('en-US',{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1}).format(n);
    default:return display(value);
  }
}

function splitCustomFormatSections(code:string):string[]|null{
 const sections:string[]=[];let start=0,quoted=false,bracketed=false;
 for(let i=0;i<code.length;i++){const c=code[i]!;if(c==='\\'){i++;continue;}if(c==='"'){if(quoted&&code[i+1]==='"'){i++;continue;}quoted=!quoted;continue;}if(!quoted&&c==='[')bracketed=true;else if(!quoted&&c===']')bracketed=false;else if(!quoted&&!bracketed&&c===';'){sections.push(code.slice(start,i));start=i+1;}}
 if(quoted||bracketed)return null;sections.push(code.slice(start));return sections.length<=4?sections:null;
}
/** Apply the fourth Excel number-format section to text values where @ is supported. */
export function formatExcelText(value:string,code:string):string|null{
 const sections=splitCustomFormatSections(code);if(!sections||sections.length<4)return null;
 const section=sections[3]??'';if(!section)return '';
 const clean=section.replace(/\[(?:Red|Blue|Green|Yellow|White|Black|Cyan|Magenta|Color\d+)\]/gi,'').replace(/\$-[0-9a-f]+/gi,'');
 if(/\[(?:[<>=].*?|[hms]+)\]/i.test(clean))return null;
 let output='';for(let i=0;i<clean.length;i++){const c=clean[i]!;if(c==='"'){let closed=false;for(i++;i<clean.length;i++){if(clean[i]==='"'){if(clean[i+1]==='"'){output+='"';i++;continue;}closed=true;break;}output+=clean[i]!;}if(!closed)return null;continue;}if(c==='\\'){if(i+1>=clean.length)return null;output+=clean[++i]!;continue;}if(c==='@'){output+=value;continue;}if(c==='_'||c==='*'){if(i+1>=clean.length)return null;i++;continue;}if(c==='['||c===']')return null;output+=c;}
 return output;
}

/** Render the common numeric subset of Excel custom formats while retaining all code in XLSX. */
export function formatExcelNumber(value:number,code:string):string|null{
 const sections=splitCustomFormatSections(code);if(!sections)return null;
 let section=sections[value<0&&sections.length>1?1:value===0&&sections.length>2?2:0]??'';if(!section)return null;
 // Numeric conditional sections need Excel's ordered condition-selection rules; preserve them in XLSX
 // and use the normal display path until that complete behavior is available.
 if(/\[(?:[<>=].*?)\]/.test(section))return null;
 const clean=section.replace(/\[(?:Red|Blue|Green|Yellow|White|Black|Cyan|Magenta|Color\d+)\]/gi,'').replace(/\$-[0-9a-f]+/gi,'').replace(/\\./g,'');
 const percent=(clean.match(/%/g)??[]).length,patternMatch=/[0#?][0#?,]*(?:\.[0#?]+)?(?:[Ee][+-][0#?]+)?/.exec(clean);if(!patternMatch)return null;const pattern=patternMatch[0],parts=pattern.split(/[Ee]/i),mantissa=parts[0]??'',integerRaw=mantissa.split('.')[0]??'',trailingScale=(integerRaw.match(/,+$/)?.[0].length??0),integerPattern=integerRaw.replace(/,+$/,''),scaled=value*(100**percent)/(1000**trailingScale);
 if(/[Ee]/.test(pattern)){const exponent=parts[1]??'',decimalCount=(mantissa.split('.')[1]??'').length;return scaled.toExponential(decimalCount).replace(/e([+-]?)(\d+)/i,(_,sign:string,digits:string)=>`E${sign||'+'}${digits.padStart(Math.max(2,exponent.replace(/\D/g,'').length), '0')}`);}
 const decimals=mantissa.split('.')[1]?.length??0,minimumDecimals=(mantissa.split('.')[1]?.match(/0/g)??[]).length,minimumInteger=(integerPattern.match(/0/g)??[]).length,formatted=new Intl.NumberFormat('en-US',{useGrouping:integerPattern.includes(','),minimumIntegerDigits:Math.max(1,Math.min(21,minimumInteger)),minimumFractionDigits:minimumDecimals,maximumFractionDigits:Math.min(20,decimals)}).format(Math.abs(scaled)),decode=(text:string)=>text.replace(/"((?:[^"]|"")*)"/g,(_,literal:string)=>literal.replaceAll('""','"'));
 const prefix=decode(clean.slice(0,patternMatch.index)),suffix=decode(clean.slice(patternMatch.index+pattern.length));
 return `${value<0&&sections.length<2?'-':''}${prefix}${formatted}${suffix}`;
}

/** Recognize date-bearing Excel number formats without treating elapsed/time-only formats as dates. */
export function isExcelDateNumberFormat(code:string):boolean{
  const elapsed=/\[(?:h+|m+|s+)\]/i.test(code),clean=code.replace(/"(?:[^"]|"")*"/g,'').replace(/\\./g,'').replace(/_.|\*./g,'').replace(/\[[^\]]+\]/g,'').replace(/am\/pm|a\/p/gi,'');
  return /[yd]/i.test(clean)||!elapsed&&/m+/i.test(clean)&&!/[hs]/i.test(clean);
}
type DateFormatToken={kind:'y'|'m'|'d'|'h'|'s'|'ampm'|'ap'|'literal';text:string;length:number};
function tokenizeExcelDateFormat(code:string):DateFormatToken[]|null{
  const tokens:DateFormatToken[]=[];
  for(let i=0;i<code.length;){
    const rest=code.slice(i),meridiem=/^(AM\/PM|A\/P)/i.exec(rest);if(meridiem){tokens.push({kind:meridiem[1]!.toUpperCase()==='AM/PM'?'ampm':'ap',text:meridiem[1]!,length:meridiem[1]!.length});i+=meridiem[1]!.length;continue;}
    const c=code[i]!;
    if(c==='$'&&/^\$-?[0-9a-f]+/i.test(rest)){const locale=/^\$-?[0-9a-f]+/i.exec(rest)![0];i+=locale.length;continue;}
    if(c==='"'){let literal='',closed=false;i++;while(i<code.length){if(code[i]==='"'){if(code[i+1]==='"'){literal+='"';i+=2;continue;}closed=true;i++;break;}literal+=code[i++]!;}if(!closed)return null;tokens.push({kind:'literal',text:literal,length:1});continue;}
    if(c==='\\'){if(i+1>=code.length)return null;tokens.push({kind:'literal',text:code[i+1]!,length:1});i+=2;continue;}
    if(c==='['){const end=code.indexOf(']',i+1);if(end<0)return null;const bracket=code.slice(i+1,end);if(/^\$-[0-9a-f]+$/i.test(bracket)){i=end+1;continue;}return null;}
    if(/[ymdhs]/i.test(c)){let end=i+1;while(end<code.length&&code[end]!.toLowerCase()===c.toLowerCase())end++;const length=end-i;if(c.toLowerCase()==='y'&&length>4||c.toLowerCase()==='m'&&length>5||c.toLowerCase()==='d'&&length>4||/[hs]/i.test(c)&&length>2)return null;tokens.push({kind:c.toLowerCase() as DateFormatToken['kind'],text:code.slice(i,end),length});i=end;continue;}
    if(/[\s\/:.,-]/.test(c)){tokens.push({kind:'literal',text:c,length:1});i++;continue;}
    return null;
  }
  return tokens;
}
/** Render the supported Excel date/time token subset; unsupported format tokens fall back safely. */
export function formatExcelDateSerial(serial:number,code:string):string|null{
  const tokens=tokenizeExcelDateFormat(code);if(!tokens||!tokens.some(token=>['y','m','d'].includes(token.kind)))return null;
  const whole=Math.floor(serial),fraction=serial-whole;if(whole<0||whole>2_958_465)return null;
  const fakeLeapDay=whole===60,date=new Date((whole<60?Date.UTC(1899,11,31):Date.UTC(1899,11,30))+whole*86_400_000),month=fakeLeapDay?2:date.getUTCMonth()+1,day=fakeLeapDay?29:date.getUTCDate(),year=fakeLeapDay?1900:date.getUTCFullYear(),milliseconds=Math.round(fraction*86_400_000),hours=Math.floor(milliseconds/3_600_000)%24,minutes=Math.floor(milliseconds/60_000)%60,seconds=Math.floor(milliseconds/1000)%60,hasMeridiem=tokens.some(token=>token.kind==='ampm'||token.kind==='ap'),monthNames=['January','February','March','April','May','June','July','August','September','October','November','December'],weekdayDate=fakeLeapDay?new Date(Date.UTC(1900,1,28)):date,weekdayNames=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  return tokens.map((token,index)=>{
    if(token.kind==='literal')return token.text;
    if(token.kind==='ampm')return hours<12?'AM':'PM';if(token.kind==='ap')return hours<12?'A':'P';
    if(token.kind==='y')return token.length===2?String(year%100).padStart(2,'0'):String(year).padStart(token.length,'0');
    if(token.kind==='d'){if(token.length===1)return String(day);if(token.length===2)return String(day).padStart(2,'0');const weekday=weekdayNames[weekdayDate.getUTCDay()]!;return token.length===3?weekday.slice(0,3):weekday;}
    if(token.kind==='h'){const hour=hasMeridiem?hours%12||12:hours;return token.length===1?String(hour):String(hour).padStart(2,'0');}
    if(token.kind==='s')return token.length===1?String(seconds):String(seconds).padStart(2,'0');
    const previous=tokens.slice(0,index).reverse().find(item=>item.kind!=='literal'),next=tokens.slice(index+1).find(item=>item.kind!=='literal'),asMinute=previous?.kind==='h'||next?.kind==='s';if(asMinute)return token.length===1?String(minutes):String(minutes).padStart(2,'0');
    if(token.length===1)return String(month);if(token.length===2)return String(month).padStart(2,'0');if(token.length===3)return monthNames[month-1]!.slice(0,3);if(token.length===4)return monthNames[month-1]!;return monthNames[month-1]!.slice(0,1);
  }).join('');
}

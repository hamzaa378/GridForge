/** GridForge native chart engine: offline, SVG-only, source-linked, no third-party chart code. */
import {Workbook} from './workbook.js';
import {parseRef,colToName} from './parser.js';
import type {TrendlineSpec} from './trendlines.js';
import {fitTrendline} from './trendlines.js';
import {formatExcelNumber} from './format.js';
export type ChartKind='column'|'bar'|'line'|'area'|'pie'|'scatter'|'doughnut'|'stacked-column'|'stacked-bar'|'percent-column'|'percent-bar'|'combo'|'radar'|'histogram'|'waterfall'|'pareto'|'box-whisker'|'funnel'|'treemap'|'stock'|'sunburst';
export const CHART_THEMES=['office','colorful','forest','sunset','ocean','monochrome'] as const;
export type ChartTheme=typeof CHART_THEMES[number];
export interface ChartErrorBars {type:'fixed'|'percentage'|'standardDeviation';value:number}
export interface ChartSeriesOptions {axis?:'primary'|'secondary';kind?:'column'|'line';color?:string;marker?:boolean;errorBars?:ChartErrorBars;trendline?:TrendlineSpec}
export type ChartNumberFormat='general'|'number'|'currency'|'percent';
export interface ChartSpec { theme?:ChartTheme; id:string; kind:ChartKind; title:string; range:string; pivotId?:string; x:number; y:number; width:number; height:number; legend?:'bottom'|'top'|'right'|'none'; xAxisTitle?:string; xAxisLabelAngle?:number; xAxisLabelInterval?:number; xAxisMin?:number; xAxisMax?:number; xAxisMajorUnit?:number; xAxisNumberFormat?:ChartNumberFormat; xAxisFormatCode?:string; yAxisTitle?:string; secondaryYAxisTitle?:string; yAxisMin?:number; yAxisMax?:number; yAxisMajorUnit?:number; secondaryYAxisMin?:number; secondaryYAxisMax?:number; secondaryYAxisMajorUnit?:number; yAxisNumberFormat?:ChartNumberFormat; secondaryYAxisNumberFormat?:ChartNumberFormat; yAxisFormatCode?:string; secondaryYAxisFormatCode?:string; doughnutHoleSize?:number; firstSliceAngle?:number; dataPointColors?:string[]; majorGridlines?:boolean; minorGridlines?:boolean; dataLabels?:boolean; dataLabelCategoryName?:boolean; dataLabelSeriesName?:boolean; dataLabelPercent?:boolean; dataLabelSeparator?:string; dataLabelPosition?:'center'|'insideEnd'|'outsideEnd'|'bestFit'; switchRowsColumns?:boolean; showFinalTotal?:boolean; waterfallTotalPoints?:number[]; histogramBins?:number; histogramBinWidth?:number; histogramUnderflow?:number; histogramOverflow?:number; seriesOptions?:ChartSeriesOptions[] }
export interface ChartSeries {name:string;values:(number|null)[]}
export interface ChartData {labels:string[];series:ChartSeries[];rows:number;range:string}
export interface ChartAxisScale {min:number;max:number}
export function calculateAxisScales(series:ChartSeries[],options:ChartSeriesOptions[]=[]):{primary:ChartAxisScale;secondary:ChartAxisScale}{
 const values=(axis:'primary'|'secondary')=>series.flatMap((item,index)=>(options[index]?.axis??'primary')===axis?item.values.filter((v):v is number=>v!==null&&Number.isFinite(v)):[]);
 const scale=(data:number[]):ChartAxisScale=>{if(!data.length)return {min:0,max:1};const low=Math.min(0,...data),high=Math.max(0,...data);return low===high?{min:low-1,max:high+1}:{min:low,max:high};};
 return {primary:scale(values('primary')),secondary:scale(values('secondary'))};
}
export const CHART_KINDS:readonly ChartKind[]=['column','bar','line','area','pie','scatter','doughnut','stacked-column','stacked-bar','percent-column','percent-bar','combo','radar','histogram','waterfall','pareto','box-whisker','funnel','treemap','stock','sunburst'];
const xml=(s:string)=>s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
export function chartRange(input:string,rows:number,cols:number){
  const match=/^\s*(\$?[A-Z]{1,3}\$?\d{1,7})\s*:\s*(\$?[A-Z]{1,3}\$?\d{1,7})\s*$/i.exec(input);
  if(!match)throw Error('Select a rectangular range such as A3:D11.');
  const a=parseRef(match[1]!.replace(/\$/g,'').toUpperCase()),b=parseRef(match[2]!.replace(/\$/g,'').toUpperCase());
  if(!a||!b||a.row>b.row||a.col>b.col||b.row>=rows||b.col>=cols)throw Error('Chart range is outside this worksheet.');
  if((b.row-a.row+1)>250||(b.col-a.col+1)>12)throw Error('Charts support up to 250 rows and 12 columns.');
  return {r1:a.row,r2:b.row,c1:a.col,c2:b.col};
}
export function validateChart(raw:unknown,rows:number,cols:number):ChartSpec {
  if(!raw||typeof raw!=='object'||Array.isArray(raw))throw Error('Invalid chart definition.');
  const c=raw as Record<string,unknown>;
  if(typeof c.id!=='string'||!/^chart-[a-zA-Z0-9_-]{1,65}$/.test(c.id))throw Error('Invalid chart identifier.');
  if(!CHART_KINDS.includes(c.kind as ChartKind))throw Error('Unsupported chart type.');
  if(typeof c.title!=='string'||c.title.length>100)throw Error('Chart title is too long.');
  if(typeof c.range!=='string')throw Error('Invalid chart source.');
  if(c.pivotId!==undefined&&(typeof c.pivotId!=='string'||!/^pivot-[A-Za-z0-9_-]{1,60}$/.test(c.pivotId)))throw Error('Invalid linked Pivot Summary identifier.');
  chartRange(c.range,rows,cols);
  for(const key of ['x','y','width','height'])if(typeof c[key]!=='number'||!Number.isFinite(c[key]))throw Error(`Invalid chart ${key}.`);
  if((c.x as number)<0||(c.x as number)>15000||(c.y as number)<0||(c.y as number)>30000||(c.width as number)<300||(c.width as number)>900||(c.height as number)<210||(c.height as number)>650)throw Error('Chart size or position is outside permitted limits.');
  if(c.theme!==undefined&&!CHART_THEMES.includes(c.theme as ChartTheme))throw Error('Invalid chart color theme.');
  if(c.legend!==undefined&&!['bottom','top','right','none'].includes(c.legend as string))throw Error('Invalid chart legend position.');
  if(c.dataLabels!==undefined&&typeof c.dataLabels!=='boolean')throw Error('Invalid chart data-label setting.');
  for(const field of ['dataLabelCategoryName','dataLabelSeriesName','dataLabelPercent'])if(c[field]!==undefined&&typeof c[field]!=='boolean')throw Error('Invalid chart data-label content setting.');
  if(c.dataLabelSeparator!==undefined&&(typeof c.dataLabelSeparator!=='string'||(c.dataLabelSeparator as string).length>20||/[\u0000-\u001f]/.test(c.dataLabelSeparator as string)))throw Error('Chart data-label separator must be text of at most 20 characters.');
  if(c.dataLabelPosition!==undefined&&!['center','insideEnd','outsideEnd','bestFit'].includes(c.dataLabelPosition as string))throw Error('Invalid chart data-label position.');
  if(c.switchRowsColumns!==undefined&&typeof c.switchRowsColumns!=='boolean')throw Error('Invalid chart data orientation.');
  if(c.showFinalTotal!==undefined&&typeof c.showFinalTotal!=='boolean')throw Error('Invalid waterfall final-total setting.');
  const rangeBounds=chartRange(c.range as string,rows,cols),sourcePointCapacity=rangeBounds.r2-rangeBounds.r1+1;
  if(c.waterfallTotalPoints!==undefined&&(!Array.isArray(c.waterfallTotalPoints)||c.waterfallTotalPoints.length>250||c.waterfallTotalPoints.some(index=>!Number.isInteger(index)||typeof index!=='number'||index<0||index>=sourcePointCapacity)))throw Error('Waterfall total point indexes must be zero-based indexes within the selected source rows.');
  if(c.histogramBins!==undefined&&(!Number.isInteger(c.histogramBins)||typeof c.histogramBins!=='number'||c.histogramBins<1||c.histogramBins>50))throw Error('Histogram bins must be a whole number from 1 to 50.');if(c.histogramBinWidth!==undefined&&(typeof c.histogramBinWidth!=='number'||!Number.isFinite(c.histogramBinWidth)||c.histogramBinWidth<=0||c.histogramBinWidth>1e15))throw Error('Histogram bin width must be a positive number no greater than 1e15.');if(c.histogramBins!==undefined&&c.histogramBinWidth!==undefined)throw Error('Choose either a histogram bin count or a bin width, not both.');for(const field of ['histogramUnderflow','histogramOverflow'])if(c[field]!==undefined&&(typeof c[field]!=='number'||!Number.isFinite(c[field])||Math.abs(c[field] as number)>1e15))throw Error('Histogram underflow and overflow thresholds must be finite numbers between -1e15 and 1e15.');if(typeof c.histogramUnderflow==='number'&&typeof c.histogramOverflow==='number'&&c.histogramUnderflow>=c.histogramOverflow)throw Error('Histogram underflow threshold must be less than the overflow threshold.');
  for(const field of ['xAxisMin','xAxisMax','yAxisMin','yAxisMax','secondaryYAxisMin','secondaryYAxisMax'])if(c[field]!==undefined&&(typeof c[field]!=='number'||!Number.isFinite(c[field])||Math.abs(c[field] as number)>1e15))throw Error('Invalid chart axis bound.');
  for(const field of ['xAxisMajorUnit','yAxisMajorUnit','secondaryYAxisMajorUnit'])if(c[field]!==undefined&&(typeof c[field]!=='number'||!Number.isFinite(c[field])||(c[field] as number)<=0||(c[field] as number)>1e15))throw Error('Chart major unit must be a positive finite number.');
  if(c.doughnutHoleSize!==undefined&&(!Number.isInteger(c.doughnutHoleSize)||typeof c.doughnutHoleSize!=='number'||c.doughnutHoleSize<10||c.doughnutHoleSize>90))throw Error('Doughnut hole size must be a whole percentage from 10 to 90.');
  if(c.firstSliceAngle!==undefined&&(!Number.isInteger(c.firstSliceAngle)||typeof c.firstSliceAngle!=='number'||c.firstSliceAngle<0||c.firstSliceAngle>360))throw Error('First slice angle must be a whole number from 0 to 360 degrees.');
  if(c.dataPointColors!==undefined&&(!['pie','doughnut'].includes(c.kind as string)||!Array.isArray(c.dataPointColors)||c.dataPointColors.length>250||c.dataPointColors.some(color=>typeof color!=='string'||!/^#[0-9a-f]{6}$/i.test(color))))throw Error('Pie and doughnut point colors must be up to 250 six-digit hexadecimal colors.');
  if(typeof c.xAxisMin==='number'&&typeof c.xAxisMax==='number'&&c.xAxisMin>=c.xAxisMax||typeof c.yAxisMin==='number'&&typeof c.yAxisMax==='number'&&c.yAxisMin>=c.yAxisMax||typeof c.secondaryYAxisMin==='number'&&typeof c.secondaryYAxisMax==='number'&&c.secondaryYAxisMin>=c.secondaryYAxisMax)throw Error('Each chart axis minimum must be less than its maximum.');
  for(const field of ['xAxisNumberFormat','yAxisNumberFormat','secondaryYAxisNumberFormat'])if(c[field]!==undefined&&!['general','number','currency','percent'].includes(c[field] as string))throw Error('Invalid chart axis number format.');
  for(const field of ['xAxisFormatCode','yAxisFormatCode','secondaryYAxisFormatCode'])if(c[field]!==undefined&&(typeof c[field]!=='string'||!(c[field] as string).length||(c[field] as string).length>100||/[\u0000-\u001f]/.test(c[field] as string)))throw Error('Custom chart axis number formats must be non-empty text of at most 100 characters.');
  for(const field of ['majorGridlines','minorGridlines'])if(c[field]!==undefined&&typeof c[field]!=='boolean')throw Error('Invalid chart gridline setting.');
  for(const field of ['xAxisTitle','yAxisTitle'])if(c[field]!==undefined&&(typeof c[field]!=='string'||(c[field] as string).length>70))throw Error('Invalid chart axis title.');
  if(c.secondaryYAxisTitle!==undefined&&(typeof c.secondaryYAxisTitle!=='string'||c.secondaryYAxisTitle.length>70))throw Error('Invalid secondary axis title.');
  if(c.xAxisLabelAngle!==undefined&&(!Number.isInteger(c.xAxisLabelAngle)||typeof c.xAxisLabelAngle!=='number'||c.xAxisLabelAngle< -90||c.xAxisLabelAngle>90))throw Error('Category axis label angle must be a whole number from -90 to 90 degrees.');
  if(c.xAxisLabelInterval!==undefined&&(!Number.isInteger(c.xAxisLabelInterval)||typeof c.xAxisLabelInterval!=='number'||c.xAxisLabelInterval<1||c.xAxisLabelInterval>250))throw Error('Category axis label interval must be a whole number from 1 to 250.');
  let seriesOptions:ChartSeriesOptions[]|undefined;
  if(c.seriesOptions!==undefined){if(!Array.isArray(c.seriesOptions)||c.seriesOptions.length>12)throw Error('Invalid chart series options.');seriesOptions=c.seriesOptions.map((raw,index)=>{if(!raw||typeof raw!=='object')throw Error('Invalid chart series options.');const o=raw as Record<string,unknown>;if(o.axis!==undefined&&!['primary','secondary'].includes(o.axis as string)||o.kind!==undefined&&!['column','line'].includes(o.kind as string)||o.color!==undefined&&(typeof o.color!=='string'||!/^#[0-9a-f]{6}$/i.test(o.color))||o.marker!==undefined&&typeof o.marker!=='boolean')throw Error('Invalid chart series options.');let errorBars:ChartErrorBars|undefined;if(o.errorBars!==undefined){const e=o.errorBars as Record<string,unknown>;if(!e||typeof e!=='object'||!['fixed','percentage','standardDeviation'].includes(e.type as string)||typeof e.value!=='number'||!Number.isFinite(e.value)||e.value<0||e.value>1e9)throw Error('Invalid chart error bars.');errorBars=e as unknown as ChartErrorBars;}let trendline:TrendlineSpec|undefined;if(o.trendline!==undefined){const t=o.trendline as Record<string,unknown>;if(!t||typeof t!=='object'||!['linear','exponential','logarithmic','polynomial','power','movingAverage'].includes(t.type as string)||!['order','period','forward','backward'].every(k=>t[k]===undefined||Number.isInteger(t[k])&&(t[k] as number)>=0&&(t[k] as number)<=100)||!['showEquation','showR2'].every(k=>t[k]===undefined||typeof t[k]==='boolean'))throw Error('Invalid chart trendline.');trendline=t as unknown as TrendlineSpec;}return { ...(o.axis?{axis:o.axis as 'primary'|'secondary'}:{}),...(o.kind?{kind:o.kind as 'column'|'line'}:{}),...(o.color?{color:o.color as string}:{}),...(o.marker!==undefined?{marker:o.marker as boolean}:{}),...(errorBars?{errorBars}:{}),...(trendline?{trendline}:{})};});}
  const waterfallTotalPoints=c.waterfallTotalPoints===undefined?undefined:[...new Set(c.waterfallTotalPoints as number[])].sort((a,b)=>a-b);
  return {...(c.theme!==undefined?{theme:c.theme as ChartTheme}:{}),id:c.id,kind:c.kind as ChartKind,title:c.title,range:c.range,...(typeof c.pivotId==='string'?{pivotId:c.pivotId}:{}),x:c.x as number,y:c.y as number,width:c.width as number,height:c.height as number,...(c.legend!==undefined?{legend:c.legend as 'bottom'|'top'|'right'|'none'}:{}),...(c.dataLabels!==undefined?{dataLabels:c.dataLabels as boolean}:{}),...(c.dataLabelCategoryName!==undefined?{dataLabelCategoryName:c.dataLabelCategoryName as boolean}:{}),...(c.dataLabelSeriesName!==undefined?{dataLabelSeriesName:c.dataLabelSeriesName as boolean}:{}),...(c.dataLabelPercent!==undefined?{dataLabelPercent:c.dataLabelPercent as boolean}:{}),...(typeof c.dataLabelSeparator==='string'?{dataLabelSeparator:c.dataLabelSeparator}:{}),...(c.dataLabelPosition!==undefined?{dataLabelPosition:c.dataLabelPosition as ChartSpec['dataLabelPosition']}:{ }),...(c.switchRowsColumns!==undefined?{switchRowsColumns:c.switchRowsColumns as boolean}:{}),...(c.showFinalTotal!==undefined?{showFinalTotal:c.showFinalTotal as boolean}:{}),...(waterfallTotalPoints!==undefined?{waterfallTotalPoints}:{}),...(c.histogramBins!==undefined?{histogramBins:c.histogramBins as number}:{}),...(c.histogramBinWidth!==undefined?{histogramBinWidth:c.histogramBinWidth as number}:{}),...(c.histogramUnderflow!==undefined?{histogramUnderflow:c.histogramUnderflow as number}:{}),...(c.histogramOverflow!==undefined?{histogramOverflow:c.histogramOverflow as number}:{}),...(c.xAxisTitle!==undefined?{xAxisTitle:c.xAxisTitle as string}:{}),...(c.xAxisLabelAngle!==undefined?{xAxisLabelAngle:c.xAxisLabelAngle as number}:{}),...(c.xAxisLabelInterval!==undefined?{xAxisLabelInterval:c.xAxisLabelInterval as number}:{}),...(c.xAxisMin!==undefined?{xAxisMin:c.xAxisMin as number}:{}),...(c.xAxisMax!==undefined?{xAxisMax:c.xAxisMax as number}:{}),...(c.xAxisMajorUnit!==undefined?{xAxisMajorUnit:c.xAxisMajorUnit as number}:{}),...(c.xAxisNumberFormat!==undefined?{xAxisNumberFormat:c.xAxisNumberFormat as ChartNumberFormat}:{}),...(typeof c.xAxisFormatCode==='string'?{xAxisFormatCode:c.xAxisFormatCode}:{}),...(c.yAxisTitle!==undefined?{yAxisTitle:c.yAxisTitle as string}:{}),...(c.secondaryYAxisTitle!==undefined?{secondaryYAxisTitle:c.secondaryYAxisTitle as string}:{}),...(c.yAxisMin!==undefined?{yAxisMin:c.yAxisMin as number}:{}),...(c.yAxisMax!==undefined?{yAxisMax:c.yAxisMax as number}:{}),...(c.yAxisMajorUnit!==undefined?{yAxisMajorUnit:c.yAxisMajorUnit as number}:{}),...(c.secondaryYAxisMin!==undefined?{secondaryYAxisMin:c.secondaryYAxisMin as number}:{}),...(c.secondaryYAxisMax!==undefined?{secondaryYAxisMax:c.secondaryYAxisMax as number}:{}),...(c.secondaryYAxisMajorUnit!==undefined?{secondaryYAxisMajorUnit:c.secondaryYAxisMajorUnit as number}:{}),...(c.yAxisNumberFormat!==undefined?{yAxisNumberFormat:c.yAxisNumberFormat as ChartNumberFormat}:{}),...(c.secondaryYAxisNumberFormat!==undefined?{secondaryYAxisNumberFormat:c.secondaryYAxisNumberFormat as ChartNumberFormat}:{}),...(typeof c.yAxisFormatCode==='string'?{yAxisFormatCode:c.yAxisFormatCode}:{}),...(typeof c.secondaryYAxisFormatCode==='string'?{secondaryYAxisFormatCode:c.secondaryYAxisFormatCode}:{}),...(c.doughnutHoleSize!==undefined?{doughnutHoleSize:c.doughnutHoleSize as number}:{}),...(c.firstSliceAngle!==undefined?{firstSliceAngle:c.firstSliceAngle as number}:{}),...(Array.isArray(c.dataPointColors)?{dataPointColors:[...(c.dataPointColors as string[])]}:{}),...(c.majorGridlines!==undefined?{majorGridlines:c.majorGridlines as boolean}:{}),...(c.minorGridlines!==undefined?{minorGridlines:c.minorGridlines as boolean}:{}),...(seriesOptions?{seriesOptions}:{})};
}
export function extractChartData(book:Workbook,range:string,switchRowsColumns=false):ChartData {
  const r=chartRange(range,book.rows,book.cols);
  const cell=(row:number,col:number)=>book.getValue(`${colToName(col)}${row+1}`);
  const scalar=(row:number,col:number):number|null=>{const v=cell(row,col);return !v.error&&typeof v.value==='number'&&Number.isFinite(v.value)?v.value:null;};
  const shown=(row:number,col:number)=>{const v=cell(row,col);return v.error?'':v.value===null?'':String(v.value);};
  const hasLabels=r.c2>r.c1&&Array.from({length:Math.min(6,r.r2-r.r1)},(_,i)=>shown(r.r1+i+1,r.c1)).some(t=>t!==''&&isNaN(Number(t)));
  const firstNumericCol=r.c1+(hasLabels?1:0);
  const heading=r.r2>r.r1&&Array.from({length:r.c2-firstNumericCol+1},(_,i)=>shown(r.r1,firstNumericCol+i)).some((t,i)=>t!==''&&scalar(r.r1,firstNumericCol+i)===null);
  const start=r.r1+(heading?1:0);
  const count=r.r2-start+1;
  const labels=Array.from({length:count},(_,i)=>hasLabels?(shown(start+i,r.c1)||String(start+i+1)):String(start+i+1));
  const series:ChartSeries[]=[];
  for(let col=firstNumericCol;col<=r.c2;col++){
    series.push({name:heading?(shown(r.r1,col)||`Series ${col-firstNumericCol+1}`):`Series ${col-firstNumericCol+1}`,
      values:Array.from({length:count},(_,i)=>scalar(start+i,col))});
  }
  if(!switchRowsColumns)return {range,labels,series,rows:count};
  return {range,labels:series.map(item=>item.name),series:labels.map((name,row)=>({name,values:series.map(item=>item.values[row]??null)})),rows:series.length};
}
export function histogramData(source:ChartData,requestedBins?:number,requestedWidth?:number,underflow?:number,overflow?:number):ChartData{
 const values=source.series.flatMap(item=>item.values.filter((v):v is number=>v!==null&&Number.isFinite(v)));
 if(!values.length)return {...source,labels:[],series:[{name:'Frequency',values:[]}],rows:0};
 if(requestedBins!==undefined&&(!Number.isInteger(requestedBins)||requestedBins<1||requestedBins>50))throw Error('Histogram bins must be a whole number from 1 to 50.');if(requestedWidth!==undefined&&(!Number.isFinite(requestedWidth)||requestedWidth<=0||requestedWidth>1e15))throw Error('Histogram bin width must be a positive number no greater than 1e15.');if(requestedBins!==undefined&&requestedWidth!==undefined)throw Error('Choose either a histogram bin count or a bin width, not both.');if(underflow!==undefined&&(!Number.isFinite(underflow)||Math.abs(underflow)>1e15)||overflow!==undefined&&(!Number.isFinite(overflow)||Math.abs(overflow)>1e15)||underflow!==undefined&&overflow!==undefined&&underflow>=overflow)throw Error('Histogram underflow threshold must be less than the overflow threshold, and thresholds must be finite.');
 const minimum=Math.min(...values),maximum=Math.max(...values),bins=requestedBins??Math.max(1,Math.min(20,Math.ceil(Math.log2(values.length)+1))),regularMin=underflow??minimum,regularMax=overflow??maximum,noInteriorRange=regularMin>regularMax,effectiveMax=noInteriorRange?regularMin:regularMax;
 const width=requestedWidth??((effectiveMax-regularMin||1)/bins),start=underflow??(requestedWidth===undefined?regularMin:Math.floor(regularMin/width)*width),calculatedBins=noInteriorRange&&requestedBins===undefined?1:requestedWidth===undefined?bins:noInteriorRange?1:Math.max(1,Math.ceil((effectiveMax-start)/width));if(!Number.isFinite(start)||!Number.isFinite(calculatedBins)||calculatedBins>50)throw Error('Histogram bin width creates more than 50 bins. Choose a larger width.');const counts=Array(calculatedBins).fill(0) as number[];
 const underflowCount=underflow===undefined?0:values.filter(value=>value<=underflow).length,overflowCount=overflow===undefined?0:values.filter(value=>value>overflow).length;
 for(const value of values){if(underflow!==undefined&&value<=underflow||overflow!==undefined&&value>overflow)continue;const index=Math.min(calculatedBins-1,Math.max(0,Math.ceil((value-start)/width)-1));counts[index]=counts[index]!+1;}
 const labels=counts.map((_,i)=>`${fmt(start+i*width)}–${fmt(start+(i+1)*width)}`),resultValues:number[]=[...counts];if(underflow!==undefined){labels.unshift(`≤ ${fmt(underflow)}`);resultValues.unshift(underflowCount);}if(overflow!==undefined){labels.push(`> ${fmt(overflow)}`);resultValues.push(overflowCount);}
 return {range:source.range,rows:resultValues.length,labels,series:[{name:'Frequency',values:resultValues}]};
}
/** Pareto input: sum each category across series, order by descending contribution, combine the tail. */
export function paretoData(source:ChartData):ChartData{
 const entries=source.labels.map((label,index)=>({label,value:source.series.reduce((total,item)=>total+(item.values[index]??0),0)}));
 if(entries.some(item=>item.value<0))throw Error('Pareto charts require non-negative values.');
 entries.sort((a,b)=>b.value-a.value||a.label.localeCompare(b.label));
 const visible=entries.slice(0,14),tail=entries.slice(14).reduce((total,item)=>total+item.value,0);if(tail>0)visible.push({label:'Other',value:tail});
 const total=visible.reduce((sum,item)=>sum+item.value,0);let running=0;
 return {range:source.range,rows:visible.length,labels:visible.map(item=>item.label),series:[{name:'Frequency',values:visible.map(item=>item.value)},{name:'Cumulative %',values:visible.map(item=>{running+=item.value;return total?running/total:0})}]};
}
/** Treemap data uses one or more hierarchy columns followed by one non-negative numeric value column. */
export function treemapData(book:Workbook,range:string):ChartData{
 const bounds=chartRange(range,book.rows,book.cols);if(bounds.c2-bounds.c1<1||bounds.r2===bounds.r1)throw Error('Treemap data needs at least one label column, one value column, and a header row.');
 const groups=new Map<string,{label:string;parent:string;value:number}>();
 for(let row=bounds.r1+1;row<=bounds.r2;row++){
  const path:string[]=[];for(let col=bounds.c1;col<bounds.c2;col++){const cell=book.getValue(`${colToName(col)}${row+1}`);if(cell.error)continue;const text=cell.value===null?'':String(cell.value).trim();if(text)path.push(text);}
  const measured=book.getValue(`${colToName(bounds.c2)}${row+1}`);if(measured.error||typeof measured.value!=='number'||!Number.isFinite(measured.value))continue;if(measured.value<0)throw Error('Hierarchy chart values must be non-negative.');if(measured.value===0||!path.length)continue;
  const label=path.join(' › '),existing=groups.get(label);if(existing)existing.value+=measured.value;else groups.set(label,{label,parent:path[0]!,value:measured.value});
 }
 const entries=[...groups.values()].sort((a,b)=>b.value-a.value||a.label.localeCompare(b.label));return {range,labels:entries.map(entry=>entry.label),series:[{name:'Value',values:entries.map(entry=>entry.value)}],rows:entries.length};
}
/** Parse Excel stock-chart inputs by header name: HLC or OHLC, each with an optional date/category column. */
export function stockData(book:Workbook,range:string):ChartData{
 const bounds=chartRange(range,book.rows,book.cols),columnCount=bounds.c2-bounds.c1+1;if(columnCount<3||columnCount>5||bounds.r1===bounds.r2)throw Error('Stock charts need a header row and High, Low and Close columns, with optional Open and date/category columns.');
 const headerText=(col:number)=>{const result=book.getValue(`${colToName(col)}${bounds.r1+1}`);return result.error||result.value===null?'':String(result.value).trim().toLowerCase();};
 const columns=new Map<string,number>();
 for(let col=bounds.c1;col<=bounds.c2;col++){const header=headerText(col);for(const field of ['open','high','low','close'])if(header===field||header.startsWith(`${field} `)||header===`${field} price`)columns.set(field,col);}
 if(['high','low','close'].some(field=>!columns.has(field))||new Set(columns.values()).size!==columns.size||![3,4].includes(columns.size)||columnCount-columns.size>1)throw Error('Name the stock columns High, Low and Close, optionally Open and a date/category column.');
 const dateCol=Array.from({length:columnCount},(_,i)=>bounds.c1+i).find(col=>![...columns.values()].includes(col));
 const labels:string[]=[],rows:(number|null)[][]=[];
 for(let row=bounds.r1+1;row<=bounds.r2;row++){
  const fields=[...(columns.has('open')?['open']:[]),'high','low','close'],values=fields.map(field=>{const result=book.getValue(`${colToName(columns.get(field)!)}${row+1}`);return !result.error&&typeof result.value==='number'&&Number.isFinite(result.value)?result.value:null;});
  if(values.every(value=>value!==null)){const high=values[fields.indexOf('high')]!,low=values[fields.indexOf('low')]!,close=values[fields.indexOf('close')]!,open=columns.has('open')?values[fields.indexOf('open')]!:undefined;if(high<Math.max(open??close,close)||low>Math.min(open??close,close)||high<low)throw Error(`Invalid stock values on row ${row+1}: High must cover Close${open===undefined?'':' and Open'}, and Low must be below them.`);}
  rows.push(values);if(dateCol!==undefined){const label=book.getValue(`${colToName(dateCol)}${row+1}`);labels.push(label.error||label.value===null?String(row+1):String(label.value));}else labels.push(String(row+1));
 }
 if(!rows.some(row=>row.every(value=>value!==null)))throw Error('Stock charts need at least one complete OHLC or HLC row.');
 const fields=[...(columns.has('open')?['Open']:[]),'High','Low','Close'];return {range,labels,series:fields.map((name,index)=>({name,values:rows.map(row=>row[index]!)})),rows:rows.length};
}
export const CHART_PALETTES:Record<ChartTheme,string[]>={
 office:['#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47','#264478','#9e480e'],
 colorful:['#2563eb','#12a594','#df8b32','#9866e6','#dc5b73','#4d97cf','#a99329','#36948a'],
 forest:['#176a52','#54a364','#c9a04c','#87bda6','#345c75','#7e9f37','#b8d8b0','#588b80'],
 sunset:['#f08b46','#db5365','#875cc0','#f4b75d','#b83467','#e17b94','#715ac1','#a34a2c'],
 ocean:['#0c7198','#34a5aa','#3f62bb','#6cc7df','#467e97','#264f7e','#74b6aa','#1d97b2'],
 monochrome:['#25374a','#50687e','#7891a8','#a9baca','#405569','#8795a3','#627489','#bdc7d0']
};
const fmt=(n:number)=>!Number.isFinite(n)?'':Math.abs(n)>=1e6?(n/1e6).toFixed(1)+'m':Math.abs(n)>=1e3?(n/1e3).toFixed(1)+'k':String(Number(n.toPrecision(4)));
const axisFmt=(n:number,format:ChartNumberFormat|undefined,customCode?:string)=>customCode?formatExcelNumber(n,customCode)??fmt(n):format==='currency'?`$${fmt(n)}`:format==='percent'?`${fmt(n*100)}%`:format==='number'?new Intl.NumberFormat('en-US',{maximumFractionDigits:2}).format(n):fmt(n);
function resolveScale(values:number[],minimum?:number,maximum?:number,includeZero=true):ChartAxisScale{
 const low=includeZero?Math.min(0,...values):Math.min(...values),high=includeZero?Math.max(0,...values):Math.max(...values),autoLow=low===high?low-1:low,autoHigh=low===high?high+1:high;
 let min=minimum??autoLow,max=maximum??autoHigh;
 if(min>=max){if(minimum!==undefined)max=min+1;else min=max-1;}
 return {min,max};
}
const esc=xml;
const tag=(x:number,y:number,content:string)=>`<text x="${x}" y="${y}" font-family="Aptos,Segoe UI,Arial,sans-serif" fill="#53677f" font-size="11">${esc(content)}</text>`;
const rotatedTag=(x:number,y:number,content:string,angle:number)=>angle?`<text x="${x}" y="${y}" transform="rotate(${angle} ${x} ${y})" font-family="Aptos,Segoe UI,Arial,sans-serif" fill="#53677f" font-size="11">${esc(content)}</text>`:tag(x,y,content);
/** Deterministic SVG for preview, floating chart and SVG download. Missing numerical cells are gaps, never zeroes. */
export function chartSvg(book:Workbook,chart:ChartSpec):string {
  let source:ChartData;try{source=chart.kind==='treemap'||chart.kind==='sunburst'?treemapData(book,chart.range):chart.kind==='stock'?stockData(book,chart.range):extractChartData(book,chart.range,chart.switchRowsColumns??false);}catch(error){const W=Math.max(300,Math.round(chart.width)),H=Math.max(210,Math.round(chart.height)),svg=`<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${xml(chart.title||chart.kind+' chart')}" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="#ffffff"/>`,title=tag(23,30,chart.title||`${chart.kind} chart`);return svg+title+tag(23,85,error instanceof Error?error.message:'Chart source data is invalid.')+'</svg>';}
  let d=source;const palette=CHART_PALETTES[chart.theme??'office'],W=Math.max(300,Math.round(chart.width)),H=Math.max(210,Math.round(chart.height));
  const svgStart=`<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(chart.title||chart.kind+' chart')}" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="#ffffff"/>`;
  const title=tag(23,30,chart.title||`${chart.kind} chart`).replace('font-size="11"','font-size="16" font-weight="700" fill="#172b4d"');
  if(chart.kind==='histogram')d=histogramData(source,chart.histogramBins,chart.histogramBinWidth,chart.histogramUnderflow,chart.histogramOverflow);
  else if(chart.kind==='pareto')try{d=paretoData(source);}catch(error){return svgStart+title+tag(23,85,error instanceof Error?error.message:'Pareto charts require non-negative values.')+'</svg>';}
  const all=d.series.flatMap(s=>s.values.filter((x):x is number=>x!==null));
  if(!d.rows||!d.series.length||!all.length)return svgStart+title+tag(23,85,'Select a range containing numeric values to plot.')+'</svg>';
  if(chart.kind==='waterfall')return waterfallSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='pareto')return paretoSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='box-whisker')return boxWhiskerSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='funnel')return funnelSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='treemap')return treemapSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='stock')return stockSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='sunburst')return sunburstSvg(chart,d,palette,W,H,svgStart,title);
  if(chart.kind==='pie'||chart.kind==='doughnut'){
    const s=d.series[0]!,pairs=s.values.map((v,i)=>({v:v??0,label:d.labels[i]??'',sourceIndex:i})).filter(x=>x.v>0).slice(0,16);
    const total=pairs.reduce((a,b)=>a+b.v,0);if(!total)return svgStart+title+tag(23,85,'Pie charts require positive values.')+'</svg>';
    const cx=Math.max(100,Math.min(W*.39,215)),cy=chart.legend==='top'?H*.62:H*.56,rad=Math.min(H*.31,W*.29,125),innerRadius=rad*(chart.doughnutHoleSize??57)/100;let angle=-Math.PI/2+(chart.firstSliceAngle??0)*Math.PI/180,out='';
    pairs.forEach(({v,label,sourceIndex},i)=>{const span=v/total*2*Math.PI;const end=angle+span;const x1=cx+rad*Math.cos(angle),y1=cy+rad*Math.sin(angle),x2=cx+rad*Math.cos(end),y2=cy+rad*Math.sin(end),sliceColor=chart.dataPointColors?.[sourceIndex]??palette[sourceIndex%palette.length]!;
      if(pairs.length===1)out+=`<circle cx="${cx}" cy="${cy}" r="${rad}" fill="${sliceColor}"/>`;
      else out+=`<path d="M ${cx} ${cy} L ${x1} ${y1} A ${rad} ${rad} 0 ${span>Math.PI?1:0} 1 ${x2} ${y2} Z" fill="${sliceColor}"/>`;
      if(chart.dataLabels||chart.dataLabelCategoryName||chart.dataLabelSeriesName||chart.dataLabelPercent){
        const mid=angle+span/2,labelParts=[...(chart.dataLabelCategoryName?[label]:[]),...(chart.dataLabelSeriesName?[s.name]:[]),...(chart.dataLabels?[fmt(v)]:[]),...(chart.dataLabelPercent?[`${(v/total*100).toFixed(1)}%`]:[])],position=chart.dataLabelPosition??'outsideEnd',outside=position==='outsideEnd'||position==='bestFit'&&span<.5,labelRadius=position==='center'?rad*.5:outside?rad*1.23:rad*.78;
        if(labelParts.length&&span>.12){const lx=cx+labelRadius*Math.cos(mid),ly=cy+labelRadius*Math.sin(mid)+4,content=labelParts.join(chart.dataLabelSeparator??' ');if(outside){const sx=cx+rad*.94*Math.cos(mid),sy=cy+rad*.94*Math.sin(mid),ex=cx+rad*1.08*Math.cos(mid),ey=cy+rad*1.08*Math.sin(mid),anchor=Math.cos(mid)<-.15?'end':Math.cos(mid)>.15?'start':'middle',tx=lx+(anchor==='end'?-3:anchor==='start'?3:0);out+=`<path d="M ${sx} ${sy} L ${ex} ${ey} L ${tx} ${ly-4}" fill="none" stroke="#68798b" stroke-width=".8"/>`+`<text x="${tx}" y="${ly}" text-anchor="${anchor}" font-family="Aptos,Segoe UI,Arial,sans-serif" fill="#334155" font-size="10">${esc(content)}</text>`;}else out+=`<text x="${lx}" y="${ly}" text-anchor="middle" font-family="Aptos,Segoe UI,Arial,sans-serif" fill="#ffffff" stroke="#26364a" stroke-width="2.4" paint-order="stroke" font-size="10" font-weight="600">${esc(content)}</text>`;}
      }
      if(i<8){if(chart.legend==='top'){const x=22+(i%3)*Math.max(90,(W-44)/3),y=52+Math.floor(i/3)*18;out+=`<rect x="${x}" y="${y-9}" width="9" height="9" fill="${palette[i%palette.length]}"/>${tag(x+13,y,`${label.slice(0,10)} ${Math.round(v/total*100)}%`)}`;}else{const y=64+i*23;out+=`<rect x="${W*.73}" y="${y-10}" width="10" height="10" fill="${palette[i%palette.length]}"/>${tag(W*.73+16,y,`${label.slice(0,14)} (${Math.round(v/total*100)}%)`)}`;}}
      angle=end;});if(chart.kind==='doughnut')out+=`<circle cx="${cx}" cy="${cy}" r="${innerRadius}" fill="white"/>`;return svgStart+title+out+'</svg>';
  }
  if(chart.kind==='radar'){
    const categories=d.labels.slice(0,32),count=categories.length,series=d.series.slice(0,6);
    if(count<3)return svgStart+title+tag(23,85,'Radar chart requires at least three data rows.')+'</svg>';
    const cx=W*.48,cy=H*.54,radius=Math.min(W*.32,H*.33),max=Math.max(1,...all.map(Math.abs));let out='';
    for(let step=1;step<=4;step++){const rr=radius*step/4;out+=`<circle cx="${cx}" cy="${cy}" r="${rr}" fill="none" stroke="#dce5ef"/>`;}
    categories.forEach((label,i)=>{const theta=-Math.PI/2+i*2*Math.PI/count,ex=cx+radius*Math.cos(theta),ey=cy+radius*Math.sin(theta);out+=`<line x1="${cx}" y1="${cy}" x2="${ex}" y2="${ey}" stroke="#dee8f1"/>${tag(cx+(radius+12)*Math.cos(theta)-17,cy+(radius+12)*Math.sin(theta)+4,label.slice(0,9))}`;});
    series.forEach((series,si)=>{const points=categories.map((_,i)=>{const v=Math.max(0,series.values[i]??0),theta=-Math.PI/2+i*2*Math.PI/count;return `${cx+radius*(v/max)*Math.cos(theta)},${cy+radius*(v/max)*Math.sin(theta)}`;}).join(' ');out+=`<polygon points="${points}" fill="${palette[si%palette.length]}" fill-opacity=".12" stroke="${palette[si%palette.length]}" stroke-width="2"/>`;});
    if(chart.legend!=='none')series.forEach((series,i)=>{const y=chart.legend==='top'?48:H-19,ty=chart.legend==='top'?57:H-10;out+=`<rect x="${18+i*112}" y="${y}" width="10" height="10" fill="${palette[i%palette.length]}"/>${tag(32+i*112,ty,series.name.slice(0,13))}`;});
    return svgStart+title+out+'</svg>';
  }
  const hasSecondary=chart.kind!=='scatter'&&(chart.seriesOptions?.some(option=>option.axis==='secondary')??false);
  const L=58,T=chart.legend==='top'?78:55,R=hasSecondary?54:22,B=(chart.legend==='none'||chart.legend==='top'?55:74)+(chart.xAxisTitle?23:0),px=L,py=T,pw=W-L-R,ph=H-T-B;
  const percent=chart.kind==='percent-column'||chart.kind==='percent-bar';
  const stacked=percent||chart.kind==='stacked-column'||chart.kind==='stacked-bar';
  const dataSeries=percent?d.series.map(series=>({name:series.name,values:series.values.map((v,i)=>{if(v===null)return null;const total=d.series.reduce((sum,s)=>sum+Math.max(0,s.values[i]??0),0);return total>0?Math.max(0,v)/total*100:0;})})):d.series;
  const stackMinMax=stacked?Array.from({length:d.rows},(_,i)=>dataSeries.reduce((tot,s)=>{const v=s.values[i]??0;if(v>=0)tot.pos+=v;else tot.neg+=v;return tot;},{pos:0,neg:0})):[];
  const isScatter=chart.kind==='scatter',values=isScatter&&d.series.length>=2?d.series.slice(1).flatMap(item=>item.values):all,scatterValues=values.filter((v):v is number=>v!==null&&Number.isFinite(v)),scatterXValues=isScatter?(d.series[0]?.values.filter((v):v is number=>v!==null&&Number.isFinite(v))??[]):[];
  const secondaryIndices=new Set((chart.seriesOptions??[]).map((option,index)=>option.axis==='secondary'?index:-1).filter(index=>index>=0));
  const primaryValues=dataSeries.flatMap((series,index)=>secondaryIndices.has(index)?[]:series.values.filter((v):v is number=>v!==null));
  const secondaryValues=dataSeries.flatMap((series,index)=>secondaryIndices.has(index)?series.values.filter((v):v is number=>v!==null):[]);
  const axisScales=calculateAxisScales(dataSeries,chart.seriesOptions);
  const rawMin=percent?0:stacked?Math.min(0,...stackMinMax.map(x=>x.neg)):hasSecondary?axisScales.primary.min:isScatter&&scatterValues.length?Math.min(...scatterValues):Math.min(0,...values.filter((v):v is number=>v!==null)),rawMax=stacked?Math.max(0,...stackMinMax.map(x=>x.pos)):hasSecondary?axisScales.primary.max:isScatter&&scatterValues.length?Math.max(...scatterValues):Math.max(0,...values.filter((v):v is number=>v!==null));
  const primaryScale=resolveScale(isScatter?scatterValues:primaryValues.length?primaryValues:values.filter((v):v is number=>v!==null),chart.yAxisMin??rawMin,chart.yAxisMax??rawMax,!isScatter),secondaryScale=resolveScale(secondaryValues,chart.secondaryYAxisMin??axisScales.secondary.min,chart.secondaryYAxisMax??axisScales.secondary.max),min=primaryScale.min,max=primaryScale.max;
  const yOf=(v:number)=>py+ph-(v-min)/(max-min)*ph;
  const secMin=secondaryScale.min,secMax=secondaryScale.max;
  const yOfSeries=(i:number,v:number)=>secondaryIndices.has(i)?py+ph-(v-secMin)/(secMax-secMin)*ph:yOf(v);
  let axisTicks:number[];
  if(chart.yAxisMajorUnit!==undefined){
    const majorUnit=Math.max(chart.yAxisMajorUnit,(max-min)/60),first=min,last=max,count=Math.max(1,Math.floor((last-first)/majorUnit+1e-12)+1);
    axisTicks=Array.from({length:count},(_,i)=>first+i*majorUnit).filter(v=>v<=last+majorUnit*1e-10);
    if(!axisTicks.length)axisTicks=[min,max];
  }else axisTicks=Array.from({length:5},(_,i)=>min+(max-min)*i/4);
  const horizontalBar=['bar','stacked-bar','percent-bar'].includes(chart.kind);let out='';
  for(let t=0;t<axisTicks.length;t++){const value=axisTicks[t]!;if(horizontalBar){const x=px+(value-min)/(max-min)*pw;if(chart.majorGridlines!==false)out+=`<line x1="${x}" x2="${x}" y1="${py}" y2="${py+ph}" stroke="#e6edf5" stroke-width="1"/>`;out+=tag(x-10,py+ph+17,axisFmt(value,chart.yAxisNumberFormat,chart.yAxisFormatCode));if(chart.minorGridlines&&t<axisTicks.length-1)for(let m=1;m<4;m++){const mx=x+(px+(axisTicks[t+1]!-min)/(max-min)*pw-x)*m/4;out+=`<line x1="${mx}" x2="${mx}" y1="${py}" y2="${py+ph}" stroke="#f1f4f8" stroke-width=".7"/>`;}}
    else{const y=yOf(value);if(chart.majorGridlines!==false)out+=`<line x1="${px}" x2="${px+pw}" y1="${y}" y2="${y}" stroke="#e6edf5" stroke-width="1"/>`;out+=tag(3,y+4,axisFmt(value,chart.yAxisNumberFormat,chart.yAxisFormatCode));if(chart.minorGridlines&&t<axisTicks.length-1)for(let m=1;m<4;m++){const my=y+(yOf(axisTicks[t+1]!)-y)*m/4;out+=`<line x1="${px}" x2="${px+pw}" y1="${my}" y2="${my}" stroke="#f1f4f8" stroke-width=".7"/>`;}}}
  const zero=Math.max(0,Math.min(1,(0-min)/(max-min))),zeroX=px+zero*pw,zeroY=py+ph-zero*ph;
  out+=horizontalBar?`<line x1="${zeroX}" x2="${zeroX}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${py+ph}" y2="${py+ph}" stroke="#9bacbf"/>`:`<line x1="${px}" x2="${px}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${zeroY}" y2="${zeroY}" stroke="#9bacbf"/>`;
  const showN=Math.min(d.rows,90),n=Math.max(1,showN),series=chart.kind==='scatter'&&d.series.length>=2?d.series.slice(1,9):dataSeries.slice(0,8);
  const xAt=(i:number)=>px+(i+.5)*pw/n;
  if(chart.kind==='bar'||chart.kind==='stacked-bar'||chart.kind==='percent-bar'){
    const block=ph/n,band=block*.78/series.length;
    const accumPos=Array(n).fill(0),accumNeg=Array(n).fill(0);
    series.forEach((s,si)=>s.values.slice(0,n).forEach((v,i)=>{if(v===null)return;const base=stacked?(v>=0?accumPos[i]:accumNeg[i]):0;if(stacked){if(v>=0)accumPos[i]+=v;else accumNeg[i]+=v;}const zeroX=px+pw*(base-min)/(max-min),endX=px+pw*(base+v-min)/(max-min);
      out+=`<rect x="${Math.min(zeroX,endX)}" y="${py+i*block+(block-(stacked?band*series.length:band*series.length))/2+(stacked?0:si*band)}" width="${Math.max(1,Math.abs(endX-zeroX))}" height="${Math.max(2,stacked?band*series.length-2:band-2)}" fill="${palette[si%palette.length]}"/>`; }));
  }else if(chart.kind==='scatter'){
    const xScale=resolveScale(scatterXValues,chart.xAxisMin,chart.xAxisMax,false),xMin=xScale.min,xMax=xScale.max,xUnit=chart.xAxisMajorUnit===undefined?undefined:Math.max(chart.xAxisMajorUnit,(xMax-xMin)/60),xTicks=xUnit===undefined?Array.from({length:5},(_,index)=>xMin+(xMax-xMin)*index/4):Array.from({length:Math.max(1,Math.floor((xMax-xMin)/xUnit+1e-12)+1)},(_,index)=>xMin+index*xUnit).filter(value=>value<=xMax+xUnit*1e-10);
    for(const value of xTicks){const x=px+(value-xMin)/(xMax-xMin)*pw;if(chart.majorGridlines!==false)out+=`<line x1="${x}" x2="${x}" y1="${py}" y2="${py+ph}" stroke="#e6edf5" stroke-width="1"/>`;out+=tag(x-14,py+ph+17,axisFmt(value,chart.xAxisNumberFormat,chart.xAxisFormatCode));}
    series.forEach((item,si)=>item.values.slice(0,n).forEach((v,i)=>{const x=d.series[0]!.values[i];if(v===null||x===null||x===undefined)return;
      const sx=px+((x-xMin)/(xMax-xMin||1))*pw,sy=yOf(v),color=chart.seriesOptions?.[si]?.color??palette[si%palette.length];if(chart.seriesOptions?.[si]?.marker!==false)out+=`<circle cx="${sx}" cy="${sy}" r="4.5" fill="${color}" opacity=".88"><title>${esc(`${item.name} · ${x}: ${v}`)}</title></circle>`;
      if((chart.dataLabels||chart.dataLabelCategoryName||chart.dataLabelSeriesName)&&d.rows<=30){const label=[...(chart.dataLabelCategoryName?[d.labels[i]??'']:[]),...(chart.dataLabelSeriesName?[item.name]:[]),...(chart.dataLabels?[fmt(v)]:[])].filter(Boolean).join(chart.dataLabelSeparator??' ');if(label)out+=tag(sx+6,sy-6,label);}
    }));
  }else{
    const segment=pw/n,barWidth=Math.max(2,segment*(chart.kind==='histogram'?.96:.7)/series.length);
    const accumPos=Array(n).fill(0),accumNeg=Array(n).fill(0);
    series.forEach((s,si)=>{
      const columnSeries=chart.kind==='combo'?(chart.seriesOptions?.[si]?.kind??(si<series.length-1?'column':'line'))==='column':(['column','stacked-column','percent-column','histogram'].includes(chart.kind)&&chart.seriesOptions?.[si]?.kind!=='line');
      if(columnSeries){s.values.slice(0,n).forEach((v,i)=>{if(v===null)return;const base=stacked?(v>=0?accumPos[i]:accumNeg[i]):0;if(stacked){if(v>=0)accumPos[i]+=v;else accumNeg[i]+=v;}const zero=yOfSeries(si,base),top=yOfSeries(si,base+v),gap=chart.kind==='histogram'?.02:.15;out+=`<rect x="${px+i*segment+segment*gap+(stacked?0:si*barWidth)}" y="${Math.min(top,zero)}" width="${stacked?barWidth*.92*series.length:barWidth*.92}" height="${Math.max(1,Math.abs(top-zero))}" rx="1" fill="${chart.seriesOptions?.[si]?.color??palette[si%palette.length]}"/>`;});}
      else{let points:string[]=[];const color=chart.seriesOptions?.[si]?.color??palette[si%palette.length],flush=()=>{if(points.length>1){const line=points.join(' ');if(chart.kind==='area')out+=`<polygon points="${points[0]!.split(',')[0]},${yOf(0)} ${line} ${points[points.length-1]!.split(',')[0]},${yOf(0)}" fill="${color}" opacity=".16"/>`;out+=`<polyline fill="none" stroke="${color}" stroke-width="2.8" stroke-linejoin="round" points="${line}"/>`;}points=[];};
        s.values.slice(0,n).forEach((v,i)=>{if(v===null){flush();return;}const x=xAt(i),y=yOfSeries(si,v);points.push(`${x},${y}`);if(chart.seriesOptions?.[si]?.marker!==false)out+=`<circle cx="${x}" cy="${y}" r="2.8" fill="${color}"/>`;});flush();}
      const bars=chart.seriesOptions?.[si]?.errorBars;if(bars){const vals=s.values.slice(0,n).filter((v):v is number=>v!==null),mean=vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0,std=vals.length?Math.sqrt(vals.reduce((a,b)=>a+(b-mean)**2,0)/vals.length):0;s.values.slice(0,n).forEach((v,i)=>{if(v===null)return;const amount=bars.type==='fixed'?bars.value:bars.type==='percentage'?Math.abs(v)*bars.value/100:std,x=xAt(i),y1=yOfSeries(si,v+amount),y2=yOfSeries(si,v-amount),color=chart.seriesOptions?.[si]?.color??palette[si%palette.length];out+=`<g stroke="${color}" stroke-width="1.2"><line x1="${x}" x2="${x}" y1="${y1}" y2="${y2}"/><line x1="${x-4}" x2="${x+4}" y1="${y1}" y2="${y1}"/><line x1="${x-4}" x2="${x+4}" y1="${y2}" y2="${y2}"/></g>`;});}
    });
    const labelStep=chart.xAxisLabelInterval??Math.max(1,Math.ceil(n/8));for(let i=0;i<n;i+=labelStep){if(horizontalBar)out+=tag(px-54,py+(i+.65)*ph/n,d.labels[i]!.slice(0,9));else out+=rotatedTag(xAt(i)-10,py+ph+21,d.labels[i]!.slice(0,10),chart.xAxisLabelAngle??0);}
  }
  if(hasSecondary){out+=`<line x1="${px+pw}" x2="${px+pw}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/>`;const unit=chart.secondaryYAxisMajorUnit===undefined?undefined:Math.max(chart.secondaryYAxisMajorUnit,(secMax-secMin)/60),first=secMin,count=unit===undefined?5:Math.max(1,Math.floor((secMax-first)/unit+1e-12)+1),ticks=unit===undefined?Array.from({length:5},(_,i)=>secMin+(secMax-secMin)*i/4):Array.from({length:count},(_,i)=>first+i*unit).filter(value=>value<=secMax+unit*1e-10);for(const value of ticks){const y=py+ph-(value-secMin)/(secMax-secMin)*ph;out+=tag(px+pw+3,y+4,axisFmt(value,chart.secondaryYAxisNumberFormat??chart.yAxisNumberFormat,chart.secondaryYAxisFormatCode??chart.yAxisFormatCode));}if(chart.secondaryYAxisTitle)out+=`<text x="${W-9}" y="${H/2}" transform="rotate(-90 ${W-9} ${H/2})" font-size="10" fill="#53677f">${esc(chart.secondaryYAxisTitle)}</text>`;}
  for(const [si,options]of (chart.seriesOptions??[]).entries())if(options.trendline){const fit=fitTrendline(dataSeries[si]?.values??[],options.trendline);if(!fit.error){const pts=fit.points.filter(p=>p.y!==null&&p.x>=1-(options.trendline!.backward??0)&&p.x<=n+(options.trendline!.forward??0)).map(p=>`${px+(p.x-.5)*pw/n},${yOfSeries(si,p.y!)}`).join(' ');if(pts)out+=`<polyline fill="none" stroke="${options.color??palette[si%palette.length]}" stroke-width="2" stroke-dasharray="6 3" points="${pts}"/>`;const note=options.trendline.showEquation?fit.equation:options.trendline.showR2&&fit.rSquared!==null?`R² = ${fit.rSquared.toFixed(4)}`:null;if(note)out+=tag(px+pw-130,py+14+si*14,note);}}
  if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-28),chart.legend==='none'?H-18:H-35,chart.xAxisTitle.slice(0,50));
  if(chart.yAxisTitle)out+=`<text x="15" y="${H/2}" transform="rotate(-90 15 ${H/2})" font-size="11" fill="#53677f">${esc(chart.yAxisTitle)}</text>`;
  if((chart.dataLabels||chart.dataLabelCategoryName||chart.dataLabelSeriesName)&&d.rows<=30&&chart.kind!=='scatter'){
    const position=chart.dataLabelPosition??'outsideEnd',labelY=(value:number,pointY:number,baseY:number)=>position==='center'?(pointY+baseY)/2+4:position==='insideEnd'?(value>=0?pointY+13:pointY-3):(value>=0?pointY-7:pointY+14);
    const labelText=(value:number,si:number,i:number)=>[...(chart.dataLabelCategoryName?[d.labels[i]??'']:[]),...(chart.dataLabelSeriesName?[series[si]?.name??'']:[]),...(chart.dataLabels?[fmt(value)]:[])].filter(Boolean).join(chart.dataLabelSeparator??' ');
    if(horizontalBar){const block=ph/n,band=block*.78/series.length,positive=Array(n).fill(0),negative=Array(n).fill(0);series.forEach((item,si)=>item.values.slice(0,n).forEach((value,i)=>{if(value===null)return;const base=stacked?(value>=0?positive[i]:negative[i]):0;if(stacked){if(value>=0)positive[i]+=value;else negative[i]+=value;}const startX=px+pw*(base-min)/(max-min),endX=px+pw*(base+value-min)/(max-min),centerY=py+i*block+(block-band*series.length)/2+(stacked?band*series.length/2:si*band+band/2),labelX=position==='center'?(startX+endX)/2:position==='insideEnd'?(value>=0?endX-22:endX+3):(value>=0?endX+3:endX-22);out+=tag(labelX,centerY+4,labelText(value,si,i));}));}
    else if(chart.kind==='stacked-column'||chart.kind==='percent-column')for(let i=0;i<n;i++){const total=series.reduce((sum,s)=>sum+Math.max(0,s.values[i]??0),0);if(total>0){const text=[...(chart.dataLabelCategoryName?[d.labels[i]??'']:[]),...(chart.dataLabelSeriesName?series.map(item=>item.name).join(', '):[]),...(chart.dataLabels?[fmt(total)]:[])].filter(Boolean).join(chart.dataLabelSeparator??' ');out+=tag(xAt(i)-12,labelY(total,yOf(total),yOf(0)),text);}}
    else series.forEach((s,si)=>s.values.slice(0,n).forEach((v,i)=>{if(v!==null){const x=(chart.kind==='column'||chart.kind==='combo')?px+(i+.1)*pw/n+si*(pw/n*.7/series.length):xAt(i)-12,isLine=chart.kind==='combo'&&chart.seriesOptions?.[si]?.kind==='line',pointY=yOfSeries(si,v),baseY=isLine?pointY:yOfSeries(si,0);out+=tag(x,labelY(v,pointY,baseY),labelText(v,si,i));}}));
  }
  if(series.length>1&&chart.legend!=='none'){series.forEach((s,i)=>{const x=L+i*Math.min(145,pw/series.length),y=chart.legend==='top'?56:H-20,ty=chart.legend==='top'?65:H-11;out+=`<rect x="${x}" y="${y}" width="10" height="10" fill="${chart.seriesOptions?.[i]?.color??palette[i%palette.length]}"/>${tag(x+15,ty,s.name.slice(0,17))}`;});}
  return svgStart+title+out+'</svg>';
}
function paretoSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const bars=(data.series[0]?.values??[]).map(value=>value??0),cumulative=(data.series[1]?.values??[]).map(value=>value??0),count=bars.length;if(!count||bars.every(value=>value===0))return svgStart+title+tag(23,85,'Pareto chart needs at least one positive value.')+'</svg>';
 const max=Math.max(...bars),L=58,T=54,R=50,B=72,px=L,py=T,pw=W-L-R,ph=H-T-B,yOf=(value:number)=>py+ph-value/Math.max(1,max)*ph,step=pw/count,barWidth=step*.68;let out='';
 for(let tick=0;tick<=4;tick++){const value=max*tick/4,y=yOf(value);if(chart.majorGridlines!==false)out+=`<line x1="${px}" x2="${px+pw}" y1="${y}" y2="${y}" stroke="#e6edf5"/>`;out+=tag(4,y+4,fmt(value));out+=tag(px+pw+4,y+4,`${tick*25}%`);}
 out+=`<line x1="${px}" x2="${px}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px+pw}" x2="${px+pw}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${py+ph}" y2="${py+ph}" stroke="#9bacbf"/>`;
 const points:{x:number;y:number}[]=[];bars.forEach((value,index)=>{const x=px+index*step+(step-barWidth)/2,y=yOf(value);out+=`<rect x="${x}" y="${y}" width="${barWidth}" height="${Math.max(1,py+ph-y)}" fill="${palette[0]}"/>`;points.push({x:px+(index+.5)*step,y:py+ph-cumulative[index]!*ph});if(chart.dataLabels&&count<=30)out+=tag(x,y-5,fmt(value));if(index%Math.max(1,Math.ceil(count/9))===0)out+=tag(x-7,py+ph+21,(data.labels[index]??'').slice(0,10));});
 if(points.length>1)out+=`<polyline fill="none" stroke="${palette[1]}" stroke-width="2.5" stroke-linejoin="round" points="${points.map(point=>`${point.x},${point.y}`).join(' ')}"/>`;
 points.forEach(({x,y},index)=>{out+=`<circle cx="${x}" cy="${y}" r="3" fill="${palette[1]}"/>`;if(chart.dataLabels&&count<=20)out+=tag(x-8,y-8,`${Math.round(cumulative[index]!*100)}%`);});
 if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-30),H-15,chart.xAxisTitle.slice(0,50));if(chart.yAxisTitle)out+=`<text x="14" y="${H/2}" transform="rotate(-90 14 ${H/2})" font-size="11" fill="#53677f">${esc(chart.yAxisTitle)}</text>`;
 out+=`<rect x="${L}" y="${H-23}" width="10" height="10" fill="${palette[0]}"/>${tag(L+14,H-14,'Frequency')}<line x1="${L+105}" x2="${L+121}" y1="${H-18}" y2="${H-18}" stroke="${palette[1]}" stroke-width="2.5"/>${tag(L+126,H-14,'Cumulative %')}`;
 return svgStart+title+out+'</svg>';
}
function boxWhiskerSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const samples=data.series.map(series=>({name:series.name,values:series.values.filter((value):value is number=>value!==null&&Number.isFinite(value)).sort((a,b)=>a-b)})).filter(sample=>sample.values.length>0);
 if(!samples.length)return svgStart+title+tag(23,85,'Box and whisker charts need numeric samples in at least one series.')+'</svg>';
 const quantile=(values:number[],p:number)=>{const position=(values.length-1)*p,lower=Math.floor(position),fraction=position-lower;return values[lower]!+(values[Math.min(values.length-1,lower+1)]!-values[lower]!)*fraction;};
 const all=samples.flatMap(sample=>sample.values),minimum=Math.min(...all),maximum=Math.max(...all),scale=resolveScale(all),L=62,T=54,R=24,B=70,px=L,py=T,pw=W-L-R,ph=H-T-B,yOf=(value:number)=>py+ph-(value-scale.min)/(scale.max-scale.min)*ph,step=pw/samples.length,boxWidth=Math.min(56,step*.48);let out='';
 for(let tick=0;tick<=4;tick++){const value=scale.min+(scale.max-scale.min)*tick/4,y=yOf(value);if(chart.majorGridlines!==false)out+=`<line x1="${px}" x2="${px+pw}" y1="${y}" y2="${y}" stroke="#e6edf5"/>`;out+=tag(3,y+4,axisFmt(value,chart.yAxisNumberFormat,chart.yAxisFormatCode));}
 out+=`<line x1="${px}" x2="${px}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${py+ph}" y2="${py+ph}" stroke="#9bacbf"/>`;
 samples.forEach((sample,index)=>{const q1=quantile(sample.values,.25),median=quantile(sample.values,.5),q3=quantile(sample.values,.75),spread=q3-q1,lowFence=q1-1.5*spread,highFence=q3+1.5*spread,low=sample.values.find(value=>value>=lowFence)!,high=[...sample.values].reverse().find(value=>value<=highFence)!,cx=px+(index+.5)*step,color=palette[index%palette.length];
  out+=`<line x1="${cx}" x2="${cx}" y1="${yOf(high)}" y2="${yOf(q3)}" stroke="${color}" stroke-width="2"/><line x1="${cx}" x2="${cx}" y1="${yOf(q1)}" y2="${yOf(low)}" stroke="${color}" stroke-width="2"/><line x1="${cx-boxWidth*.28}" x2="${cx+boxWidth*.28}" y1="${yOf(high)}" y2="${yOf(high)}" stroke="${color}" stroke-width="2"/><line x1="${cx-boxWidth*.28}" x2="${cx+boxWidth*.28}" y1="${yOf(low)}" y2="${yOf(low)}" stroke="${color}" stroke-width="2"/><rect x="${cx-boxWidth/2}" y="${yOf(q3)}" width="${boxWidth}" height="${Math.max(1,yOf(q1)-yOf(q3))}" fill="${color}" fill-opacity=".3" stroke="${color}" stroke-width="2"/><line x1="${cx-boxWidth/2}" x2="${cx+boxWidth/2}" y1="${yOf(median)}" y2="${yOf(median)}" stroke="${color}" stroke-width="3"/>`;
  sample.values.filter(value=>value<lowFence||value>highFence).forEach(value=>{out+=`<circle cx="${cx}" cy="${yOf(value)}" r="3" fill="white" stroke="${color}" stroke-width="1.5"/>`;});out+=tag(cx-28,py+ph+21,sample.name.slice(0,12));
 });
 if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-30),H-15,chart.xAxisTitle.slice(0,50));if(chart.yAxisTitle)out+=`<text x="14" y="${H/2}" transform="rotate(-90 14 ${H/2})" font-size="11" fill="#53677f">${esc(chart.yAxisTitle)}</text>`;
 return svgStart+title+out+'</svg>';
}
function funnelSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const values=data.series[0]?.values??[],items=data.labels.map((label,index)=>({label,value:values[index]??null})).filter((item):item is {label:string;value:number}=>item.value!==null&&Number.isFinite(item.value));
 if(items.some(item=>item.value<0))return svgStart+title+tag(23,85,'Funnel charts require non-negative values.')+'</svg>';
 if(!items.length||items.every(item=>item.value===0))return svgStart+title+tag(23,85,'Funnel charts need at least one positive value.')+'</svg>';
 const visible=items.slice(0,30),L=75,T=52,R=38,B=44,px=L,py=T,pw=W-L-R,ph=H-T-B,step=ph/visible.length,max=Math.max(...visible.map(item=>item.value)),center=px+pw/2,halfWidth=pw*.46;let out='';
 visible.forEach((item,index)=>{const next=visible[index+1]?.value??0,top=halfWidth*item.value/max,bottom=halfWidth*next/max,y1=py+index*step,y2=y1+step,color=palette[index%palette.length];out+=`<path d="M ${center-top} ${y1} L ${center+top} ${y1} L ${center+bottom} ${y2} L ${center-bottom} ${y2} Z" fill="${color}" stroke="#ffffff" stroke-width="1.5"/>`;if(step>=22)out+=tag(Math.max(px+2,center-top+6),y1+Math.min(step-4,16),`${item.label.slice(0,14)}  ${fmt(item.value)}`);});
 if(items.length>visible.length)out+=tag(px,py+ph+17,`${items.length-visible.length} additional stages not shown`);if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-30),H-14,chart.xAxisTitle.slice(0,50));
 return svgStart+title+out+'</svg>';
}
function treemapSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const entries=data.labels.map((label,index)=>({label,parent:label.split(' › ')[0]??'',value:data.series[0]?.values[index]??0})).filter(item=>item.value>0);
 if(!entries.length)return svgStart+title+tag(23,85,'Treemap needs positive values and category labels.')+'</svg>';
 const L=16,T=48,R=16,B=20,plot={x:L,y:T,w:W-L-R,h:H-T-B},total=entries.reduce((sum,item)=>sum+item.value,0);let remaining={...plot},row:typeof entries=[];const rects:{item:typeof entries[number];x:number;y:number;w:number;h:number}[]=[];
 const worst=(items:typeof entries,side:number)=>{if(!items.length||side<=0)return Infinity;const areas=items.map(item=>item.value/total*plot.w*plot.h),sum=areas.reduce((a,b)=>a+b,0),max=Math.max(...areas),min=Math.min(...areas);return Math.max(side*side*max/(sum*sum),sum*sum/(side*side*min));};
 const layoutRow=(items:typeof entries,box:{x:number;y:number;w:number;h:number})=>{const area=items.reduce((sum,item)=>sum+item.value/total*plot.w*plot.h,0);if(box.w>=box.h){const band=area/box.h;let x=box.x;for(const item of items){const w=(item.value/total*plot.w*plot.h)/box.h;rects.push({item,x,y:box.y,w,h:box.h});x+=w;}return {...box,x:box.x+band,w:Math.max(0,box.w-band)};}const band=area/box.w;let y=box.y;for(const item of items){const h=(item.value/total*plot.w*plot.h)/box.w;rects.push({item,x:box.x,y,w:box.w,h});y+=h;}return {...box,y:box.y+band,h:Math.max(0,box.h-band)};};
 for(const item of entries){const side=Math.min(remaining.w,remaining.h);if(row.length&&worst([...row,item],side)>worst(row,side)){remaining=layoutRow(row,remaining);row=[];}row.push(item);}if(row.length)layoutRow(row,remaining);
 const parentOrder=[...new Set(entries.map(item=>item.parent))],colorOf=(parent:string)=>palette[parentOrder.indexOf(parent)%palette.length]!;let out='';for(const cell of rects){const {item,x,y,w,h}=cell,color=colorOf(item.parent),pad=1;out+=`<rect x="${x+pad}" y="${y+pad}" width="${Math.max(0,w-2*pad)}" height="${Math.max(0,h-2*pad)}" rx="2" fill="${color}" fill-opacity=".82" stroke="#ffffff" stroke-width="1"/>`;if(w>50&&h>24)out+=tag(x+5,y+Math.min(17,h-5),item.label.slice(0,Math.max(4,Math.floor(w/7))));if(chart.dataLabels&&w>48&&h>38)out+=tag(x+5,y+Math.min(33,h-5),fmt(item.value));}
 return svgStart+title+out+'</svg>';
}
function stockSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const count=data.rows;if(!count)return svgStart+title+tag(23,85,'Stock chart needs at least one complete OHLC row.')+'</svg>';
 const hasOpen=data.series[0]?.name==='Open',high=data.series[hasOpen?1:0]!.values,low=data.series[hasOpen?2:1]!.values,close=data.series[hasOpen?3:2]!.values,open=hasOpen?data.series[0]!.values:close,all=[...high,...low].filter((value):value is number=>value!==null),rawMin=Math.min(...all),rawMax=Math.max(...all),pad=rawMin===rawMax?Math.max(1,Math.abs(rawMin)*.05):0,min=rawMin-pad,max=rawMax+pad,L=64,T=52,R=18,B=66,px=L,py=T,pw=W-L-R,ph=H-T-B,scale=max-min||1,yOf=(value:number)=>py+ph-(value-min)/scale*ph,step=pw/count,bodyWidth=Math.max(2,Math.min(15,step*.58));let out='';
 for(let tick=0;tick<=4;tick++){const value=min+(max-min)*tick/4,y=yOf(value);if(chart.majorGridlines!==false)out+=`<line x1="${px}" x2="${px+pw}" y1="${y}" y2="${y}" stroke="#e6edf5"/>`;out+=tag(3,y+4,axisFmt(value,chart.yAxisNumberFormat,chart.yAxisFormatCode));}
 out+=`<line x1="${px}" x2="${px}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${py+ph}" y2="${py+ph}" stroke="#9bacbf"/>`;
 for(let index=0;index<count;index++){const cx=px+(index+.5)*step,h=high[index],l=low[index],c=close[index],o=open[index];if(h===null||h===undefined||l===null||l===undefined||c===null||c===undefined||o===null||o===undefined)continue;const up=c>=o,color=hasOpen?(up?'#18865b':'#d44949'):palette[0]!,openY=yOf(o),closeY=yOf(c),top=Math.min(openY,closeY),height=Math.max(2,Math.abs(openY-closeY));out+=`<line x1="${cx}" x2="${cx}" y1="${yOf(h)}" y2="${yOf(l)}" stroke="${color}" stroke-width="1.5"/>`;if(hasOpen)out+=`<rect x="${cx-bodyWidth/2}" y="${top}" width="${bodyWidth}" height="${height}" fill="${color}" stroke="${color}" stroke-width="1"/>`;else out+=`<line x1="${cx}" x2="${cx+bodyWidth/2}" y1="${closeY}" y2="${closeY}" stroke="${color}" stroke-width="1.5"/>`;if(chart.dataLabels&&count<=40)out+=tag(cx-10,top-5,fmt(c));if(index%Math.max(1,Math.ceil(count/9))===0)out+=tag(cx-18,py+ph+21,(data.labels[index]??'').slice(0,12));}
 if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-30),H-14,chart.xAxisTitle.slice(0,50));if(chart.yAxisTitle)out+=`<text x="14" y="${H/2}" transform="rotate(-90 14 ${H/2})" font-size="11" fill="#53677f">${esc(chart.yAxisTitle)}</text>`;
 return svgStart+title+out+'</svg>';
}
function sunburstSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 type Node={name:string;value:number;children:Map<string,Node>};const root:Node={name:'',value:0,children:new Map()};
 data.labels.forEach((path,index)=>{const value=data.series[0]?.values[index]??0;if(value<=0)return;let parent=root;for(const name of path.split(' › ')){let child=parent.children.get(name);if(!child){child={name,value:0,children:new Map()};parent.children.set(name,child);}child.value+=value;parent=child;}});
 const children=[...root.children.values()].sort((a,b)=>b.value-a.value),total=children.reduce((sum,node)=>sum+node.value,0);if(!children.length||!total)return svgStart+title+tag(23,85,'Sunburst needs positive values and hierarchy labels.')+'</svg>';
 let depth=0;const measure=(nodes:Node[],level:number)=>{depth=Math.max(depth,level);for(const node of nodes)measure([...node.children.values()],level+1);};measure(children,0);depth=Math.max(1,depth);
 const cx=W*.43,cy=H*.56,radius=Math.min(W*.38,H*.39,185),ring=radius/depth;let out='';
 const point=(r:number,a:number)=>({x:cx+r*Math.cos(a),y:cy+r*Math.sin(a)});
 const sector=(r0:number,r1:number,start:number,end:number)=>{const span=end-start;if(span>=Math.PI*2-1e-7){if(r0===0)return `<circle cx="${cx}" cy="${cy}" r="${r1}"`;return `M ${cx-r1} ${cy} A ${r1} ${r1} 0 1 1 ${cx+r1} ${cy} A ${r1} ${r1} 0 1 1 ${cx-r1} ${cy} Z M ${cx-r0} ${cy} A ${r0} ${r0} 0 1 0 ${cx+r0} ${cy} A ${r0} ${r0} 0 1 0 ${cx-r0} ${cy} Z`;};const outerStart=point(r1,start),outerEnd=point(r1,end),large=span>Math.PI?1:0;if(r0===0)return `M ${cx} ${cy} L ${outerStart.x} ${outerStart.y} A ${r1} ${r1} 0 ${large} 1 ${outerEnd.x} ${outerEnd.y} Z`;const innerEnd=point(r0,end),innerStart=point(r0,start);return `M ${outerStart.x} ${outerStart.y} A ${r1} ${r1} 0 ${large} 1 ${outerEnd.x} ${outerEnd.y} L ${innerEnd.x} ${innerEnd.y} A ${r0} ${r0} 0 ${large} 0 ${innerStart.x} ${innerStart.y} Z`;};
 const roots=children;const walk=(nodes:Node[],parentValue:number,start:number,end:number,level:number,rootIndex:number,path:string[])=>{let cursor=start;for(const node of nodes){const span=(end-start)*node.value/parentValue,next=cursor+span,r0=level*ring,r1=(level+1)*ring,color=palette[(rootIndex+level*2)%palette.length]!,d=sector(r0,r1,cursor,next),labelPath=[...path,node.name].join(' › ');if(d.startsWith('<circle'))out+=`${d} fill="${color}" stroke="#fff" stroke-width="1.5"/>`;else out+=`<path d="${d}" fill="${color}" fill-opacity="${level===0?'.88':String(Math.max(.48,.82-level*.08))}" fill-rule="evenodd" stroke="#ffffff" stroke-width="1.5"><title>${esc(labelPath)}: ${fmt(node.value)}</title></path>`;if(span>.23&&r1-r0>16){const mid=(cursor+next)/2,pt=point((r0+r1)/2,mid);out+=`<text x="${pt.x}" y="${pt.y}" text-anchor="middle" dominant-baseline="middle" font-family="Aptos,Segoe UI,Arial,sans-serif" font-size="9" fill="#172b4d">${esc(node.name.slice(0,14))}</text>`;}const descendants=[...node.children.values()].sort((a,b)=>b.value-a.value);if(descendants.length)walk(descendants,node.value,cursor,next,level+1,rootIndex,[...path,node.name]);cursor=next;}};
 walk(roots,total,-Math.PI/2,Math.PI*1.5,0,0,[]);
 out+=`<circle cx="${cx}" cy="${cy}" r="${Math.max(8,ring*.38)}" fill="#ffffff" stroke="#e2e8f0"/>`;
 if(chart.legend!=='none')roots.slice(0,10).forEach((node,index)=>{const x=W*.77,y=62+index*21;out+=`<rect x="${x}" y="${y-9}" width="10" height="10" fill="${palette[index%palette.length]}"/>${tag(x+15,y,node.name.slice(0,15))}`;});
 return svgStart+title+out+'</svg>';
}
export function waterfallTotalPointIndexes(chart:ChartSpec,pointCount:number):number[]{
 const indexes=new Set(chart.waterfallTotalPoints??[]);if(chart.showFinalTotal===true&&pointCount>0)indexes.add(pointCount-1);
 for(const index of indexes)if(!Number.isInteger(index)||index<0||index>=pointCount)throw Error('Waterfall total point indexes must be within the selected source data rows.');
 return [...indexes].sort((a,b)=>a-b);
}
function waterfallSvg(chart:ChartSpec,data:ChartData,palette:string[],W:number,H:number,svgStart:string,title:string):string{
 const deltas=data.series[0]?.values??[],count=Math.min(data.rows,80),points:{start:number;end:number;delta:number|null;label:string}[]=[];let total=0;
 const totalIndexes=waterfallTotalPointIndexes(chart,data.rows);for(const index of totalIndexes)if(deltas[index]===null||deltas[index]===undefined)throw Error(`Waterfall total at item ${index+1} must have a numeric value.`);const totals=new Set(totalIndexes.filter(index=>index<count));
 for(let i=0;i<count;i++){const delta=deltas[i]??null,isTotal=totals.has(i),start=isTotal?0:total,end=delta===null?total:isTotal?delta:total+delta;if(delta!==null)total=end;points.push({start,end,delta,label:data.labels[i]??String(i+1)});}
 if(!points.some(point=>point.delta!==null))return svgStart+title+tag(23,85,'Waterfall chart needs at least one numeric value.')+'</svg>';
 const lows=points.flatMap(point=>[point.start,point.end]),min=Math.min(0,...lows),max=Math.max(0,...lows),scale=resolveScale(lows,min,max),L=60,T=56,R=22,B=74,px=L,py=T,pw=W-L-R,ph=H-T-B,yOf=(value:number)=>py+ph-(value-scale.min)/(scale.max-scale.min)*ph,segment=pw/Math.max(1,count),barWidth=Math.max(5,segment*.62),zero=yOf(0);let out='';
 for(let tick=0;tick<=4;tick++){const value=scale.min+(scale.max-scale.min)*tick/4,y=yOf(value);if(chart.majorGridlines!==false)out+=`<line x1="${px}" x2="${px+pw}" y1="${y}" y2="${y}" stroke="#e6edf5"/>`;out+=tag(4,y+4,axisFmt(value,chart.yAxisNumberFormat,chart.yAxisFormatCode));}
 out+=`<line x1="${px}" x2="${px+pw}" y1="${zero}" y2="${zero}" stroke="#8394a8"/><line x1="${px}" x2="${px}" y1="${py}" y2="${py+ph}" stroke="#9bacbf"/><line x1="${px}" x2="${px+pw}" y1="${py+ph}" y2="${py+ph}" stroke="#9bacbf"/>`;
 points.forEach((point,index)=>{const x=px+index*segment+(segment-barWidth)/2;if(index>0){const prior=points[index-1]!;out+=`<line x1="${px+(index-1)*segment+(segment+barWidth)/2}" x2="${x}" y1="${yOf(prior.end)}" y2="${yOf(prior.end)}" stroke="#8b9aab" stroke-dasharray="3 2"/>`;}
  if(point.delta!==null){const color=totals.has(index)?'#4472c4':point.delta>=0?'#18865b':'#d45d47',top=yOf(Math.max(point.start,point.end)),bottom=yOf(Math.min(point.start,point.end));out+=`<rect x="${x}" y="${top}" width="${barWidth}" height="${Math.max(1,bottom-top)}" rx="1" fill="${color}"/>`;if(chart.dataLabels)out+=tag(x-6,top-5,fmt(point.delta));}
  if(index%Math.max(1,Math.ceil(count/9))===0)out+=tag(x-8,py+ph+21,point.label.slice(0,10));
 });
 if(chart.xAxisTitle)out+=tag(Math.max(70,W/2-30),H-15,chart.xAxisTitle.slice(0,50));if(chart.yAxisTitle)out+=`<text x="14" y="${H/2}" transform="rotate(-90 14 ${H/2})" font-size="11" fill="#53677f">${esc(chart.yAxisTitle)}</text>`;
 return svgStart+title+out+'</svg>';
}

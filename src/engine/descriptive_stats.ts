import {Workbook,type Edit} from './workbook.js';
import {parseRef,colToName} from './parser.js';

export interface DescriptiveStatsResult {edits:Edit[];outputRange:string;count:number;mode:number|null}

/** Build a compact Excel ToolPak-style descriptive statistics report as ordinary cells. */
export function buildDescriptiveStatistics(book:Workbook,sourceRange:string,destination:string):DescriptiveStatsResult{
 const parseRange=(text:string)=>{const parts=text.trim().replace(/\$/g,'').toUpperCase().split(':');if(parts.length>2)throw Error('Enter one rectangular cell range.');const first=parseRef(parts[0]??''),last=parseRef(parts[1]??parts[0]??'');if(!first||!last||first.row>last.row||first.col>last.col||last.row>=book.rows||last.col>=book.cols)throw Error('Enter a valid ascending cell range inside this worksheet.');return {r1:first.row,r2:last.row,c1:first.col,c2:last.col};};
 const source=parseRange(sourceRange),target=parseRange(destination);if(target.r1!==target.r2||target.c1!==target.c2)throw Error('Choose one destination cell for the report.');
 const cells=(source.r2-source.r1+1)*(source.c2-source.c1+1);if(cells>12000)throw Error('Descriptive Statistics supports selections up to 12,000 cells.');
 const values:number[]=[];for(let row=source.r1;row<=source.r2;row++)for(let col=source.c1;col<=source.c2;col++){const value=book.getValue(`${colToName(col)}${row+1}`);if(!value.error&&typeof value.value==='number'&&Number.isFinite(value.value))values.push(value.value);}
 if(!values.length)throw Error('The selected range contains no numeric values.');
 const sorted=[...values].sort((a,b)=>a-b),n=values.length,mean=values.reduce((sum,value)=>sum+value,0)/n,median=n%2?sorted[(n-1)/2]!:(sorted[n/2-1]!+sorted[n/2]!)/2;
 const counts=new Map<number,number>();for(const value of values)counts.set(value,(counts.get(value)??0)+1);const highest=Math.max(...counts.values()),mode=highest>1?Math.min(...[...counts].filter(([,count])=>count===highest).map(([value])=>value)):null;
 const sumSquares=values.reduce((sum,value)=>sum+(value-mean)**2,0),sampleVariance=n>1?sumSquares/(n-1):null,sampleDeviation=sampleVariance===null?null:Math.sqrt(sampleVariance),standardError=sampleDeviation===null?null:sampleDeviation/Math.sqrt(n);
 let skewness:number|null=null,kurtosis:number|null=null;if(n>=3&&sampleDeviation){skewness=n/((n-1)*(n-2))*values.reduce((sum,value)=>sum+((value-mean)/sampleDeviation)**3,0);}if(n>=4&&sampleVariance){kurtosis=n*(n+1)/((n-1)*(n-2)*(n-3))*values.reduce((sum,value)=>sum+((value-mean)**4)/(sampleVariance**2),0)-3*(n-1)**2/((n-2)*(n-3));}
 const rows:[string,number|null|'#N/A'][]=[['Mean',mean],['Standard Error',standardError??'#N/A'],['Median',median],['Mode',mode??'#N/A'],['Standard Deviation (Sample)',sampleDeviation??'#N/A'],['Sample Variance',sampleVariance??'#N/A'],['Kurtosis',kurtosis??'#N/A'],['Skewness',skewness??'#N/A'],['Range',sorted[n-1]!-sorted[0]!],['Minimum',sorted[0]!],['Maximum',sorted[n-1]!],['Sum',values.reduce((sum,value)=>sum+value,0)],['Count',n]];
 if(target.c1+1>=book.cols||target.r1+rows.length>=book.rows)throw Error('The report needs two columns and 14 rows of space at its destination.');
 const output={r1:target.r1,r2:target.r1+rows.length,c1:target.c1,c2:target.c1+1};if(!(output.r2<source.r1||source.r2<output.r1||output.c2<source.c1||source.c2<output.c1))throw Error('Place the report outside the selected source range.');
 const edits:Edit[]=[{ref:`${colToName(target.c1)}${target.r1+1}`,raw:'Statistic'},{ref:`${colToName(target.c1+1)}${target.r1+1}`,raw:'Value'}];for(let index=0;index<rows.length;index++){const row=rows[index]!,rowNumber=target.r1+index+2;edits.push({ref:`${colToName(target.c1)}${rowNumber}`,raw:row[0]},{ref:`${colToName(target.c1+1)}${rowNumber}`,raw:typeof row[1]==='number'?String(row[1]):'=#N/A'});}
 return {edits,outputRange:`${colToName(output.c1)}${output.r1+1}:${colToName(output.c2)}${output.r2+1}`,count:n,mode};
}

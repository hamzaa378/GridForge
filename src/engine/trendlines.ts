export type TrendlineType='linear'|'exponential'|'logarithmic'|'polynomial'|'power'|'movingAverage';
export interface TrendlineSpec {type:TrendlineType;order?:number;period?:number;forward?:number;backward?:number;showEquation?:boolean;showR2?:boolean}
export interface TrendlineResult {points:{x:number;y:number|null}[];equation:string|null;rSquared:number|null;error?:string}
const solve=(a:number[][],b:number[]):number[]|null=>{
 const n=b.length,m=a.map((row,i)=>[...row,b[i]!]);
 for(let c=0;c<n;c++){let p=c;for(let r=c+1;r<n;r++)if(Math.abs(m[r]![c]!)>Math.abs(m[p]![c]!))p=r;
  if(Math.abs(m[p]![c]!)<1e-12)return null;[m[c],m[p]]=[m[p]!,m[c]!];const d=m[c]![c]!;for(let j=c;j<=n;j++)m[c]![j]=m[c]![j]!/d;
  for(let r=0;r<n;r++)if(r!==c){const f=m[r]![c]!;for(let j=c;j<=n;j++)m[r]![j]=m[r]![j]!-f*m[c]![j]!;}
 }
 return m.map(row=>row[n]!);
};
const linearFit=(xs:number[],ys:number[])=>{const n=xs.length,sx=xs.reduce((a,b)=>a+b,0),sy=ys.reduce((a,b)=>a+b,0),sxx=xs.reduce((a,x)=>a+x*x,0),sxy=xs.reduce((a,x,i)=>a+x*ys[i]!,0),d=n*sxx-sx*sx;if(Math.abs(d)<1e-12)return null;const slope=(n*sxy-sx*sy)/d,intercept=(sy-slope*sx)/n;return {slope,intercept,predict:(x:number)=>slope*x+intercept};};
export function fitTrendline(values:(number|null)[],spec:TrendlineSpec,xValues?:(number|null)[]):TrendlineResult{
 const points=values.map((y,i)=>({x:xValues?xValues[i]:i+1,y})).filter((p):p is {x:number;y:number}=>typeof p.y==='number'&&Number.isFinite(p.y)&&typeof p.x==='number'&&Number.isFinite(p.x)).sort((a,b)=>a.x-b.x);
 const period=Math.max(2,Math.min(100,Math.trunc(spec.period??2))),order=Math.max(2,Math.min(6,Math.trunc(spec.order??2))),forward=Math.max(0,Math.min(50,Math.trunc(spec.forward??0))),backward=Math.max(0,Math.min(50,Math.trunc(spec.backward??0)));
 if(spec.type==='movingAverage'){
  if(points.length<period)return {points:[],equation:null,rSquared:null,error:'Moving-average period exceeds the available data.'};
  const step=points.length>1?(points.at(-1)!.x-points[0]!.x)/(points.length-1):1,averages=points.map((_,i)=>i+1<period?null:points.slice(i-period+1,i+1).reduce((s,x)=>s+x.y,0)/period),first=averages.find((v):v is number=>v!==null)!,last=averages.at(-1)!;
  const out:{x:number;y:number|null}[]=[];for(let i=-backward;i<points.length+forward;i++){const x=points[0]!.x+i*step,y=i<0?first:i>=points.length?last:averages[i]??null;out.push({x,y});}
  return {points:out,equation:null,rSquared:null};
 }
 let predict:((x:number)=>number|null)|null=null,equation:string|null=null;
 if(spec.type==='linear'){const fit=linearFit(points.map(p=>p.x),points.map(p=>p.y));if(fit){predict=fit.predict;equation=`y = ${fit.slope.toPrecision(5)}x ${fit.intercept<0?'−':'+'} ${Math.abs(fit.intercept).toPrecision(5)}`;}}
 else if(spec.type==='exponential'||spec.type==='power'){
  const transformed=points.filter(p=>p.y>0&&(spec.type!=='power'||p.x>0)),fit=linearFit(transformed.map(p=>spec.type==='power'?Math.log(p.x):p.x),transformed.map(p=>Math.log(p.y)));
  if(fit){predict=(x)=>spec.type==='power'&&x<=0?null:Math.exp(fit.intercept+fit.slope*(spec.type==='power'?Math.log(x):x));equation=spec.type==='power'?`y = ${Math.exp(fit.intercept).toPrecision(5)}x^${fit.slope.toPrecision(5)}`:`y = ${Math.exp(fit.intercept).toPrecision(5)}e^(${fit.slope.toPrecision(5)}x)`;if(transformed.length<points.length)return {points:[],equation:null,rSquared:null,error:`${spec.type} trendlines require ${spec.type==='power'?'positive x and y':'positive y'} values.`};}
 }
 else if(spec.type==='logarithmic'){
  const valid=points.filter(p=>p.x>0),fit=linearFit(valid.map(p=>Math.log(p.x)),valid.map(p=>p.y));if(valid.length<points.length)return {points:[],equation:null,rSquared:null,error:'Logarithmic trendlines require positive x values.'};if(fit){predict=x=>x<=0?null:fit.predict(Math.log(x));equation=`y = ${fit.slope.toPrecision(5)} ln(x) ${fit.intercept<0?'−':'+'} ${Math.abs(fit.intercept).toPrecision(5)}`;}
 }
 else if(spec.type==='polynomial'){
  if(points.length<=order)return {points:[],equation:null,rSquared:null,error:'Polynomial order must be lower than the number of valid points.'};
  const matrix=Array.from({length:order+1},(_,i)=>Array.from({length:order+1},(_,j)=>points.reduce((s,p)=>s+p.x**(i+j),0))),rhs=Array.from({length:order+1},(_,i)=>points.reduce((s,p)=>s+p.y*p.x**i,0)),coefs=solve(matrix,rhs);
  if(coefs){predict=x=>coefs.reduce((s,c,i)=>s+c*x**i,0);equation=`y = ${coefs.map((c,i)=>`${c.toPrecision(4)}x${i?`^${i}`:''}`).reverse().join(' + ')}`;}
 }
 if(!predict)return {points:[],equation:null,rSquared:null,error:'The data is degenerate for this trendline type.'};
 const mean=points.reduce((s,p)=>s+p.y,0)/points.length,ssTot=points.reduce((s,p)=>s+(p.y-mean)**2,0),ssRes=points.reduce((s,p)=>s+(p.y-(predict(p.x)??NaN))**2,0),rSquared=ssTot===0?null:1-ssRes/ssTot;
 const step=points.length>1?(points.at(-1)!.x-points[0]!.x)/(points.length-1):1,lo=points[0]!.x-step*backward,hi=points.at(-1)!.x+step*forward;
 return {points:Array.from({length:Math.max(1,Math.round((hi-lo)/(step||1))+1)},(_,i)=>{const x=lo+i*(step||1),y=predict!(x);return {x,y:y!==null&&Number.isFinite(y)?y:null};}),equation,rSquared};
}

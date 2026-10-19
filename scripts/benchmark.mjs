/** Reproducible, hardware-dependent benchmark; run `npm run build && node scripts/benchmark.mjs`. */
import {performance} from 'node:perf_hooks';
import {Workbook} from '../dist/engine/workbook.js';
import {colToName} from '../dist/engine/parser.js';
const w=new Workbook(400,52),edits=[];
for(let row=1;row<=400;row++){
  edits.push({ref:`A${row}`,raw:String(row)});
  for(let col=1;col<=25;col++){
    const ref=`${colToName(col)}${row}`;
    edits.push({ref,raw:`=$A${row}+${col}`});
  }
}
const t=performance.now();const initial=w.applyEdits(edits,false);const elapsed=performance.now()-t;
const independent=w.setCell('A2','2000');
const noDependency=w.setCell('AZ400','999');
const out={node:process.version,rows:400,columns:52,capacity:20800,populated:initial.totalPopulated,formulas:400*25,
    initialBuildMs:Number(elapsed.toFixed(3)),initialRecalculated:initial.recalculated,
    singleSourceEdit:{affected:independent.affected.length,recalculated:independent.recalculated,elapsedMs:independent.durationMs},
    isolatedEdit:{affected:noDependency.affected.length,recalculated:noDependency.recalculated,elapsedMs:noDependency.durationMs}};
console.log(JSON.stringify(out,null,2));
if(independent.affected.length!==26||noDependency.affected.length!==1)process.exitCode=1;

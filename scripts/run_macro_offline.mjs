#!/usr/bin/env node
/** GridForge offline macro CLI. Uses the identical tested TypeScript executor as the app.
 * Node 20+ required for command-line developer use; the packaged desktop UI does not require Node.
 */
import {readFile, writeFile, rename, unlink, stat} from 'node:fs/promises';
import {resolve,dirname,basename,join} from 'node:path';
import {randomUUID} from 'node:crypto';
import {SpreadsheetDocument} from '../dist/engine/document.js';
import {runActionMacro,validateMacro} from '../dist/engine/macros.js';
function options(argv){
  const out={};for(let i=0;i<argv.length;i++){
    if(['--dry-run','--help'].includes(argv[i])){out[argv[i].slice(2)]=true;continue;}
    const key=argv[i];if(!['--workbook','--macro','--output','--parameters'].includes(key)||!argv[i+1])throw Error(`Unknown/missing argument: ${key}`);
    if(out[key.slice(2)])throw Error(`Duplicate ${key}`);out[key.slice(2)]=argv[++i];
  }return out;
}
const usage='Usage: node scripts/run_macro_offline.mjs --workbook workbook.json --macro macro.json [--output updated.json] [--parameters values.json] [--dry-run]';
async function readJSON(path){const file=resolve(path);if((await stat(file)).size>24*1024*1024)throw Error('Input JSON larger than 24 MB.');return JSON.parse(await readFile(file,'utf8'));}
try{
  const opts=options(process.argv.slice(2));if(opts.help){console.log(usage);process.exit(0);}
  if(!opts.workbook||!opts.macro)throw Error(usage);
  if(!opts['dry-run']&&!opts.output)throw Error('Output path is required unless --dry-run.');
  const input=resolve(opts.workbook),output=opts.output?resolve(opts.output):'';
  if(output===input)throw Error('Refusing to overwrite the input workbook; choose a new output path.');
  if(output){try{await stat(output);throw Error('Output file already exists. Choose a new path.');}catch(e){if(e.code!=='ENOENT')throw e;}}
  const doc=SpreadsheetDocument.fromJSON(await readJSON(opts.workbook));
  const macro=validateMacro(await readJSON(opts.macro));
  const parameters=opts.parameters?await readJSON(opts.parameters):undefined;
  const preview=runActionMacro(doc,macro,{dryRun:true,parameters});
  if(opts['dry-run']){console.log(JSON.stringify({status:'PREVIEW',macro:macro.name,...preview},null,2));process.exit(0);}
  const result=runActionMacro(doc,macro,{parameters});
  const temporary=join(dirname(output),`.${basename(output)}.${randomUUID()}.tmp`);
  try{await writeFile(temporary,JSON.stringify(doc.toJSON(),null,2)+'\n',{encoding:'utf8',flag:'wx'});await rename(temporary,output);}catch(e){await unlink(temporary).catch(()=>{});throw e;}
  console.log(JSON.stringify({status:'SAVED',output,macro:macro.name,values:result.values,styles:result.styles,sheets:result.sheets},null,2));
}catch(e){console.error(`MACRO ERROR: ${e?.message??String(e)}`);process.exitCode=1;}

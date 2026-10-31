import {cp, mkdir} from 'node:fs/promises';
import {build} from 'esbuild';
await mkdir('dist',{recursive:true});
await cp('public/index.html','dist/index.html');
await cp('public/styles.css','dist/styles.css');
await build({entryPoints:['node_modules/@xmldom/xmldom/lib/index.js'],bundle:true,format:'esm',platform:'neutral',outfile:'dist/vendor/xmldom.js'});
console.log('Static assets copied to dist/.');

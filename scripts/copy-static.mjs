import {cp, mkdir} from 'node:fs/promises';
await mkdir('dist',{recursive:true});
await cp('public/index.html','dist/index.html');
await cp('public/styles.css','dist/styles.css');
console.log('Static assets copied to dist/.');

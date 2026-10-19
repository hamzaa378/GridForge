import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
const customPort=process.env.PORT!==undefined;
let PORT=Number(process.env.PORT??4174);
if(!Number.isInteger(PORT)||PORT<1||PORT>65535){console.error('PORT must be an integer between 1 and 65535.');process.exit(1);}
const MIME={'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'text/javascript; charset=utf-8','.json':'application/json; charset=utf-8','.svg':'image/svg+xml'};
const server=createServer(async(req,res)=>{
  try{
    const pathname=decodeURI((req.url??'/').split('?')[0]??'/');
    const rel=pathname==='/'?'index.html':pathname.replace(/^\/+/, '');
    const full=path.resolve('dist',rel);
    if(!full.startsWith(path.resolve('dist')+path.sep)){res.writeHead(403).end('Forbidden');return;}
    const bytes=await readFile(full);
    res.writeHead(200,{'Content-Type':MIME[path.extname(full)]??'application/octet-stream','Cache-Control':'no-cache','X-Content-Type-Options':'nosniff'});res.end(bytes);
  }catch{res.writeHead(404).end('Not found');}
});
server.on('error',e=>{
  if(e.code==='EADDRINUSE'&&!customPort&&PORT<4184){
    console.log(`Port ${PORT} is in use; trying ${PORT+1}...`);
    PORT++;server.listen(PORT);return;
  }
  console.error(`Could not start GridForge on port ${PORT}: ${e.message}`);
  if(e.code==='EADDRINUSE')console.error('Use another port: $env:PORT="4190" (PowerShell), then npm run dev.');
  process.exitCode=1;
});
server.listen(PORT,()=>console.log(`GridForge running at http://localhost:${PORT}`));

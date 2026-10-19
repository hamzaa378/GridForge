import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import path from 'node:path';
const PORT=Number(process.env.PORT??4174);
const MIME={'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'text/javascript; charset=utf-8','.json':'application/json; charset=utf-8','.svg':'image/svg+xml'};
createServer(async(req,res)=>{
  try{
    const pathname=decodeURI((req.url??'/').split('?')[0]??'/');
    const rel=pathname==='/'?'index.html':pathname.replace(/^\/+/, '');
    const full=path.resolve('dist',rel);
    if(!full.startsWith(path.resolve('dist')+path.sep)){res.writeHead(403).end('Forbidden');return;}
    const bytes=await readFile(full);
    res.writeHead(200,{'Content-Type':MIME[path.extname(full)]??'application/octet-stream','Cache-Control':'no-cache','X-Content-Type-Options':'nosniff'});res.end(bytes);
  }catch{res.writeHead(404).end('Not found');}
}).listen(PORT,()=>console.log(`GridForge running at http://localhost:${PORT}`));

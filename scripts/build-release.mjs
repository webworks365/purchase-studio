import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),source=path.join(root,'public-site'),target=path.join(root,'site-release');
const version=process.env.RELEASE_SHA||process.argv[2];if(!/^[a-f0-9]{7,64}$/.test(version||''))throw Error('A release commit SHA is required');
const samples=JSON.parse(await fs.readFile(path.join(source,'samples.json'),'utf8'));if(samples.files?.length)throw Error('Private templates must never be published');
await fs.mkdir(target,{recursive:true});await fs.cp(source,target,{recursive:true});
for(const entry of await fs.readdir(target,{withFileTypes:true}))if(entry.isFile()&&/\.m?js$/.test(entry.name)){
 const file=path.join(target,entry.name),text=await fs.readFile(file,'utf8');
 await fs.writeFile(file,text.replace(/(['"])(\.\/[^'"\s]+?\.m?js)\1/g,(_,quote,ref)=>quote+ref+'?v='+version+quote));
}
const html=path.join(target,'index.html');let text=await fs.readFile(html,'utf8');text=text.replaceAll('__BUILD_VERSION__',version).replace(/((?:src|href)="[^"?]+\.(?:js|css))"/g,'$1?v='+version+'"');await fs.writeFile(html,text);
await fs.writeFile(path.join(target,'version.json'),JSON.stringify({version}));
console.log('Prepared versioned website release '+version);

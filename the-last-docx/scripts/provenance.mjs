import { readFile,readdir,writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { join,resolve,relative,sep } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url)),source=resolve(process.argv[2]??join(root,'..','docx分区'));
const hashes={},changes=[];
async function scan(path,base){for(const entry of await readdir(path,{withFileTypes:true})){const file=join(path,entry.name);if(entry.isDirectory())await scan(file,base);else if(entry.isFile()&&!entry.name.endsWith('.pyc')){
  const key=relative(base,file).split(sep).join('/'),bytes=await readFile(file),hash=createHash('sha256').update(bytes).digest('hex');hashes[key]=hash;
  const original=join(source,key.replace(/^modules\//,''));let sourceHash;
  try{sourceHash=createHash('sha256').update(await readFile(original)).digest('hex');}catch{sourceHash=null;}
  if(sourceHash!==hash)changes.push(key);
}}}
for(const entry of await readdir(join(root,'modules'),{withFileTypes:true}))if(entry.isDirectory())await scan(join(root,'modules',entry.name,'src'),root);
await writeFile(join(root,'modules.lock.json'),JSON.stringify({schema:'docx-module-provenance/v1',sourceRootHint:'../docx分区',scope:'Module src files only; package names and tsconfig changes are integration adapters.',allModuleSourcesMatch:changes.length===0,changedSources:changes,sha256:hashes},null,2)+'\n');
console.log(JSON.stringify({files:Object.keys(hashes).length,changedSources:changes}));if(changes.length)process.exitCode=1;

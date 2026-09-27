import { readFile,readdir,writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { join,resolve,relative,sep } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url)),source=resolve(process.argv[2]??join(root,'..','..','docx分区'));
const hashes={},changes=[];
const localExtensions=['modules/pptx-office','modules/xlsx-office','modules/pdf-office'];
const localExtensionFiles=[
  'modules/docx-edit/src/domain/docx-edit.ts',
  'modules/docx-edit/src/engine/adapter.ts',
  'modules/docx-edit/src/engine/table-xml.ts',
  'modules/docx-render/src/engine/adapter.ts',
];
async function scan(path,base){for(const entry of await readdir(path,{withFileTypes:true})){const file=join(path,entry.name);if(entry.isDirectory())await scan(file,base);else if(entry.isFile()&&!entry.name.endsWith('.pyc')){
  const key=relative(base,file).split(sep).join('/'),bytes=await readFile(file),hash=createHash('sha256').update(bytes).digest('hex');hashes[key]=hash;
  const original=join(source,key.replace(/^modules\//,''));let sourceHash;
  try{sourceHash=createHash('sha256').update(await readFile(original)).digest('hex');}catch{sourceHash=null;}
  if(sourceHash!==hash&&!localExtensions.some(prefix=>key.startsWith(`${prefix}/`))&&!localExtensionFiles.includes(key))changes.push(key);
}}}
for(const entry of await readdir(join(root,'modules'),{withFileTypes:true}))if(entry.isDirectory())await scan(join(root,'modules',entry.name,'src'),root);
await writeFile(join(root,'modules.lock.json'),JSON.stringify({schema:'docx-module-provenance/v1',sourceRootHint:relative(root,source).split(sep).join('/'),scope:'DOCX module src files are compared to sourceRootHint; DSH-only modules and explicitly enumerated DSH integration files are recorded as local extensions.',localExtensionModules:localExtensions,localExtensionFiles,allModuleSourcesMatch:changes.length===0,changedSources:changes,sha256:hashes},null,2)+'\n');
console.log(JSON.stringify({files:Object.keys(hashes).length,changedSources:changes}));if(changes.length)process.exitCode=1;

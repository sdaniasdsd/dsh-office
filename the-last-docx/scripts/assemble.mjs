import { cp, mkdir, readFile, writeFile, readdir } from 'node:fs/promises';
import { resolve, join, basename, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=resolve(fileURLToPath(new URL('..',import.meta.url)));
const source=resolve(process.argv[2]??join(root,'..','docx分区'));
const names=['docx-inspect','docx-easy-parse','docx-parse','docx-complex-parse','docx-create','docx-styles','docx-edit','docx-render','docx-artifact'];
for(const name of names){
  const destination=join(root,'modules',name);
  await mkdir(destination,{recursive:true});
  await cp(join(source,name),destination,{recursive:true,filter:path=>!relative(join(source,name),path).split(sep).some(p=>['node_modules','.git','_bench','_generated','__pycache__','coverage'].includes(p))});
  const path=join(destination,'package.json'),pkg=JSON.parse(await readFile(path,'utf8'));
  // Packaging-only alias; source capability ID and all exported behavior stay unchanged.
  if(name==='docx-easy-parse')pkg.name='@dsh-office-profile/docx-easy-parse';
  pkg.peerDependencies={...pkg.peerDependencies,'office-core':'*','office-safety':'*','office-files':'*','office-test-kit':'*'};
  pkg.peerDependenciesMeta={...pkg.peerDependenciesMeta,...Object.fromEntries(['office-core','office-safety','office-files','office-test-kit'].map(k=>[k,{optional:true}]))};
  for(const field of ['dependencies','devDependencies'])for(const [key,value]of Object.entries(pkg[field]??{}))if(String(value).startsWith('file:../docx-'))pkg[field][key]='*';
  await writeFile(path,JSON.stringify(pkg,null,2)+'\n');
  const configPath=join(destination,'tsconfig.json'),config=JSON.parse(await readFile(configPath,'utf8'));
  config.include=['src/**/*.ts','tests/**/*.ts','fixtures/**/*.ts'];config.exclude=['types','node_modules'];
  await writeFile(configPath,JSON.stringify(config,null,2)+'\n');
}
// Promote the EXISTING canonical public type declarations into one package per owner.
// All copies keep their old type stubs for provenance, but compilation no longer includes them.
const shared=await readFile(join(source,'docx-inspect/types/office-deps.d.ts'),'utf8');
for(const [,name,body]of shared.matchAll(/declare module '([^']+)' \{([\s\S]*?)\r?\n\}/g)){
  const directory=join(root,'packages',name);await mkdir(directory,{recursive:true});
  let contents=body;
  if(name==='office-safety')contents=contents.replace(/(export interface SafetyPolicy \{)/,'$1\n    maxBlocks?: number;\n    maxTableCells?: number;');
  await writeFile(join(directory,'index.ts'),'// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.\n'+contents+'\n');
  await writeFile(join(directory,'package.json'),JSON.stringify({name,version:'0.1.0',private:true,type:'module',types:'index.ts',exports:{'.':'./index.ts'}},null,2)+'\n');
}
console.log(`Assembled ${names.length} modules in ${root}. No source module was modified.`);

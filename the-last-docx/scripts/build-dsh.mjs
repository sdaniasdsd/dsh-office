import { build } from 'esbuild';
import { readFile,writeFile,mkdir,cp,readdir,rm } from 'node:fs/promises';
import { resolve,join,relative,dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url));
const out=resolve(root,'dist','dsh-docx');
await mkdir(join(out,'lib'),{recursive:true});
const bundledPackages=new Map();
// Preserve each Python adapter's sibling-script lookup after bundling. Source files stay untouched.
const engines={name:'python-engine-assets',setup(builder){builder.onLoad({filter:/[\\/]engine[\\/]adapter\.ts$/},async args=>{
  const match=args.path.replaceAll('\\','/').match(/\/modules\/([^/]+)\/src\/engine\/adapter\.ts$/);
  if(!match)return;
  let contents=await readFile(args.path,'utf8');
  if(!contents.includes('import.meta.url'))return;
  const name=match[1];
  contents=contents.replaceAll('import.meta.url',`new URL('./engines/${name}/adapter.js', import.meta.url).href`);
  await mkdir(join(out,'lib','engines',name),{recursive:true});
  for(const entry of await readdir(dirname(args.path)))if(entry.endsWith('.py'))await cp(join(dirname(args.path),entry),join(out,'lib','engines',name,entry));
  return {contents,loader:'ts',resolveDir:dirname(args.path)};
});}};
for(const entry of ['server','index'])await build({entryPoints:[join(root,'src',`${entry}.ts`)],outfile:join(out,'lib',`${entry}.mjs`),bundle:true,platform:'node',format:'esm',target:'node22',plugins:[engines],banner:{js:"import { createRequire as __createRequire } from 'node:module'; const require = __createRequire(import.meta.url);"},metafile:true}).then(async result=>{
  await writeFile(join(out,'lib',`${entry}.build.json`),JSON.stringify(result.metafile,null,2));
  for(const input of Object.keys(result.metafile.inputs)){
    let directory=dirname(resolve(root,input));
    while(directory!==root&&directory.startsWith(root)){
      if(directory.split(/[\\/]/).includes('node_modules')){
        try{
          const packageJson=JSON.parse(await readFile(join(directory,'package.json'),'utf8'));
          if(packageJson.name)bundledPackages.set(packageJson.name,directory);
          break;
        }catch{}
      }
      const parent=dirname(directory);if(parent===directory)break;directory=parent;
    }
  }
});
const licenseRoot=join(out,'third-party-licenses');
for(const [packageName,packageRoot] of bundledPackages){
  const licenseFiles=(await readdir(packageRoot,{withFileTypes:true})).filter(entry=>entry.isFile()&&/^(license|licence|notice|copying)/i.test(entry.name));
  if(!licenseFiles.length)continue;
  const destination=join(licenseRoot,packageName.replaceAll('/','__'));
  await mkdir(destination,{recursive:true});
  for(const license of licenseFiles)await cp(join(packageRoot,license.name),join(destination,license.name));
}
await Promise.all(['server','index'].map(entry=>rm(join(out,'lib',`${entry}.build.json`),{force:true})));
await cp(join(root,'dsh'),join(out,'dsh'),{recursive:true});
await cp(join(root,'dsh','cordis.patch.yml'),join(out,'cordis.patch.yml'));
await cp(join(root,'runtime'),join(out,'runtime'),{recursive:true});
for(const name of ['ARCHITECTURE.md','THIRD_PARTY.md','modules.lock.json'])await cp(join(root,name),join(out,name));
const manifest={name:'@deepseek-ai/dsh-docx',version:'0.5.0',type:'module',license:'UNLICENSED',description:'Nine DOCX modules with dual IR, style authoring, editing, rendering and versioned delivery for DSH.',main:'dsh/index.mjs',exports:{'.':'./dsh/index.mjs','./core':'./lib/index.mjs','./package.json':'./package.json'},files:['dsh','lib','runtime','third-party-licenses','cordis.patch.yml','README.md','ARCHITECTURE.md','THIRD_PARTY.md','modules.lock.json'],engines:{node:'^22.19.0 || >=24.0.0'},os:['win32'],cpu:['x64'],dsh:{bundle:{patch:'./cordis.patch.yml'}},dependencies:{'@deepseek-ai/dsh-mcp-client':'>=0.1.0-rc.8 <1'},peerDependencies:{'@deepseek-ai/cordis':'^4.0.1'}};
await writeFile(join(out,'package.json'),JSON.stringify(manifest,null,2)+'\n');
await cp(join(root,'dsh','README.md'),join(out,'README.md'));
console.log(`DSH package built: ${out}`);

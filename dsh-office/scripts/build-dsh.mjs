import { build } from 'esbuild';
import { readFile,writeFile,mkdir,cp,readdir,rm,rename,stat,access } from 'node:fs/promises';
import { resolve,join,relative,dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url));
const VERSION='0.9.0';
const platformDir='win32-x64';
const coreOut=resolve(root,'dist','dsh-docx');
const runtimeOut=resolve(root,'dist','dsh-docx-runtime');
const fullOut=resolve(root,'dist','dsh-docx-full');
const exists=path=>access(path).then(()=>true,()=>false);
const parsed=new Map(process.argv.slice(2).filter(a=>a.startsWith('--')).map(a=>a.replace(/^--/,'').split('=')));
const explicitRuntime=parsed.get('runtime');

// ---------------------------------------------------------------------------
// 1. 核心包：JS bundle + dsh 清单 + 文档，**不含运行时**（运行时是另一个包）
// ---------------------------------------------------------------------------
await mkdir(join(coreOut,'lib'),{recursive:true});
const bundledPackages=new Map();
// Preserve each Python adapter's sibling-script lookup after bundling. Source files stay untouched.
const engines={name:'python-engine-assets',setup(builder){
  builder.onLoad({filter:/[\\/]engine[\\/]adapter\.ts$/},async args=>{
    const match=args.path.replaceAll('\\','/').match(/\/modules\/([^/]+)\/src\/engine\/adapter\.ts$/);
    if(!match)return;
    let contents=await readFile(args.path,'utf8');
    if(!contents.includes('import.meta.url'))return;
    const name=match[1];
    contents=contents.replaceAll('import.meta.url',`new URL('./engines/${name}/adapter.js', import.meta.url).href`);
    await mkdir(join(coreOut,'lib','engines',name),{recursive:true});
    for(const entry of await readdir(dirname(args.path)))if(entry.endsWith('.py'))await cp(join(dirname(args.path),entry),join(coreOut,'lib','engines',name,entry));
    return {contents,loader:'ts',resolveDir:dirname(args.path)};
  });
  builder.onLoad({filter:/[\\/]modules[\\/](pptx-office|xlsx-office)[\\/]src[\\/]index\.ts$/},async args=>{
    const match=args.path.replaceAll('\\','/').match(/\/modules\/(pptx-office|xlsx-office)\/src\/index\.ts$/);
    if(!match)return;
    const name=match[1],engineDir=join(dirname(args.path),'engine');
    const contents=(await readFile(args.path,'utf8')).replaceAll('import.meta.url',`new URL('./engines/${name}/index.js', import.meta.url).href`);
    await cp(engineDir,join(coreOut,'lib','engines',name,'engine'),{recursive:true});
    return {contents,loader:'ts',resolveDir:dirname(args.path)};
  });
}};
for(const entry of ['server','index'])await build({entryPoints:[join(root,'src',`${entry}.ts`)],outfile:join(coreOut,'lib',`${entry}.mjs`),bundle:true,platform:'node',format:'esm',target:'node22',external:['pdfjs-dist'],plugins:[engines],define:{__DSH_DOCX_VERSION__:JSON.stringify(VERSION)},banner:{js:"import { createRequire as __createRequire } from 'node:module'; const require = __createRequire(import.meta.url);"},metafile:true}).then(async result=>{
  await writeFile(join(coreOut,'lib',`${entry}.build.json`),JSON.stringify(result.metafile,null,2));
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
const licenseRoot=join(coreOut,'third-party-licenses');
for(const [packageName,packageRoot] of bundledPackages){
  const licenseFiles=(await readdir(packageRoot,{withFileTypes:true})).filter(entry=>entry.isFile()&&/^(license|licence|notice|copying)/i.test(entry.name));
  if(!licenseFiles.length)continue;
  const destination=join(licenseRoot,packageName.replaceAll('/','__'));
  await mkdir(destination,{recursive:true});
  for(const license of licenseFiles)await cp(join(packageRoot,license.name),join(destination,license.name));
}
await Promise.all(['server','index'].map(entry=>rm(join(coreOut,'lib',`${entry}.build.json`),{force:true})));
await cp(join(root,'dsh'),join(coreOut,'dsh'),{recursive:true});
await cp(join(root,'dsh','cordis.patch.yml'),join(coreOut,'cordis.patch.yml'));
for(const name of ['ARCHITECTURE.md','OFFICE_ENGINE_DECISIONS.md','THIRD_PARTY.md','modules.lock.json'])await cp(join(root,name),join(coreOut,name));
// 核心包不再携带运行时，也不再限定平台：Windows 运行时是另一个包的事。
const coreManifest={name:'@deepseek-ai/dsh-docx',version:VERSION,type:'module',license:'UNLICENSED',description:'DOCX, PPTX, XLSX and native PDF modules with parsing, editing, rendering and versioned delivery for DSH. Runtime (Python/LibreOffice/Poppler) ships separately as @deepseek-ai/dsh-docx-runtime; without it the runtime-dependent capabilities report themselves unavailable instead of failing startup.',main:'dsh/index.mjs',exports:{'.':'./dsh/index.mjs','./core':'./lib/index.mjs','./package.json':'./package.json'},files:['dsh','lib','third-party-licenses','cordis.patch.yml','README.md','ARCHITECTURE.md','OFFICE_ENGINE_DECISIONS.md','THIRD_PARTY.md','modules.lock.json'],engines:{node:'^22.19.0 || >=24.0.0'},dsh:{bundle:{patch:'./cordis.patch.yml'}},dependencies:{'@deepseek-ai/dsh-mcp-client':'>=0.1.0-rc.8 <1','pdfjs-dist':'6.3.289'},peerDependencies:{'@deepseek-ai/cordis':'^4.0.1'}};
await writeFile(join(coreOut,'package.json'),JSON.stringify(coreManifest,null,2)+'\n');
await cp(join(root,'dsh','README.md'),join(coreOut,'README.md'));
console.log(`core package built:    ${coreOut}`);

// ---------------------------------------------------------------------------
// 2. 运行时包：只装三件二进制 + 一份可校验的清单
// ---------------------------------------------------------------------------
const runtimeCandidates=[explicitRuntime,join(root,'runtime',platformDir),join(coreOut,'runtime',platformDir)].filter(Boolean);
let runtimeSource=null;
for(const candidate of runtimeCandidates){if(await exists(join(candidate,'python','python.exe'))){runtimeSource=resolve(candidate);break;}}
if(!runtimeSource){
  await rm(fullOut,{recursive:true,force:true});
  console.warn(`runtime package skipped: no runtime at ${runtimeCandidates.join(' | ')}. Run scripts/fetch-dsh-runtime.ps1 (or pass --runtime=<dir>) before packing it.`);
  console.log(`core manifest version: ${VERSION}`);
  process.exit(0);
}
await rm(runtimeOut,{recursive:true,force:true});
await mkdir(join(runtimeOut,'runtime'),{recursive:true});
const runtimeDestination=join(runtimeOut,'runtime',platformDir);
if(runtimeSource===join(coreOut,'runtime',platformDir)){
  // 旧布局里 core 自带一份运行时：直接改名搬过去，别复制 1.6 GB。
  await rename(runtimeSource,runtimeDestination);
  await rm(join(coreOut,'runtime'),{recursive:true,force:true});
}else{
  await cp(runtimeSource,runtimeDestination,{recursive:true});
}
const measure=async path=>{let bytes=0,files=0;for(const entry of await readdir(path,{withFileTypes:true})){const child=join(path,entry.name);if(entry.isDirectory()){const inner=await measure(child);bytes+=inner.bytes;files+=inner.files;}else if(entry.isFile()){const info=await stat(child);bytes+=info.size;files+=1;}}return {bytes,files};};
const components={};
for(const [name,relativePath] of Object.entries({python:join('python','python.exe'),libreoffice:join('libreoffice','program','soffice.com'),poppler:join('poppler','poppler-26.09.0','Library','bin','pdftoppm.exe')})){
  const path=join(runtimeDestination,relativePath);
  const present=await exists(path);
  const here=await measure(join(runtimeDestination,name));
  components[name]={present,entry:present?relativePath:null,bytes:here.bytes,files:here.files};
}
const runtimeManifest={schema:'dsh-office-runtime/v1',version:VERSION,platform:platformDir,components,
  note:'Core package finds this by walking up node_modules for @deepseek-ai/dsh-docx-runtime and reading runtime.json; explicit runtimeRoot/DSH_OFFICE_RUNTIME_ROOT overrides it.'};
await writeFile(join(runtimeOut,'runtime.json'),JSON.stringify(runtimeManifest,null,2)+'\n');
const runtimePackage={name:'@deepseek-ai/dsh-docx-runtime',version:VERSION,type:'module',license:'UNLICENSED',description:'Bundled offline runtime for the DSH office plugin: private Python 3.13, LibreOffice and Poppler for Windows x64. No install scripts, no downloads at run time.',files:['runtime','runtime.json','README.md'],os:['win32'],cpu:['x64']};
await writeFile(join(runtimeOut,'package.json'),JSON.stringify(runtimePackage,null,2)+'\n');
await writeFile(join(runtimeOut,'README.md'),`# @deepseek-ai/dsh-docx-runtime ${VERSION}

给 DSH office 插件用的离线运行时（Windows x64）：私有 Python 3.13（含 lxml / python-pptx / Pillow / XlsxWriter）、
LibreOffice、Poppler。没有安装脚本，也不在运行时下载任何东西。

装上它，核心包 \`@deepseek-ai/dsh-docx\` 会自己找到（向上找 \`node_modules/@deepseek-ai/dsh-docx-runtime\`
并读 \`runtime.json\`）；也可以用 \`runtimeRoot\` 配置或 \`DSH_OFFICE_RUNTIME_ROOT\` 环境变量指向别处的一份运行时。
没有它时核心包照常启动，只是依赖运行时的能力（解析、PPTX、渲染）会在调用时明确报不可用。

组件与体积（见 \`runtime.json\`）：${Object.entries(components).map(([name,info])=>`${name} ${info.present?'✓':'✗'} ${(info.bytes/1048576).toFixed(0)}MB/${info.files} files`).join('，')}
`);
console.log(`runtime package built: ${runtimeOut}  (source ${runtimeSource})`);

// ---------------------------------------------------------------------------
// 3. 元包：一步装好 = 核心 + 运行时
// ---------------------------------------------------------------------------
await mkdir(fullOut,{recursive:true});
const fullPackage={name:'@deepseek-ai/dsh-docx-full',version:VERSION,type:'module',license:'UNLICENSED',description:'Meta package: the DSH office plugin core plus its Windows x64 runtime, for a one-step install.',private:false,dependencies:{'@deepseek-ai/dsh-docx':VERSION,'@deepseek-ai/dsh-docx-runtime':VERSION}};
await writeFile(join(fullOut,'package.json'),JSON.stringify(fullPackage,null,2)+'\n');
await writeFile(join(fullOut,'README.md'),`# @deepseek-ai/dsh-docx-full ${VERSION}

只写依赖的元包：\`@deepseek-ai/dsh-docx\`（核心，约 15 MB）+ \`@deepseek-ai/dsh-docx-runtime\`（Windows x64 运行时，约 1.6 GB）。

想一步装好就用它；想省体积、或已经有自己的一份 LibreOffice/Python，就只装核心再指定
\`runtimeRoot\` / \`DSH_OFFICE_RUNTIME_ROOT\`。两个子包版本号与它保持一致。
`);
console.log(`full package built:    ${fullOut}`);

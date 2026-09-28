import { build } from 'esbuild';
import { readFile,writeFile,mkdir,cp,readdir,rm,rename,stat,access } from 'node:fs/promises';
import { resolve,join,relative,dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url));
const VERSION='0.10.2';
// `dsh-mcp-client` loads this peer set while evaluating the DSH plugin entry.
// Pin the family together so a profile install never falls back to an unrelated
// copy in npm's cache (which can make the harness crash-loop during startup).
const DSH_MCP_VERSION='0.1.0-rc.8';
const DSH_MCP_STARTUP_DEPENDENCIES=Object.freeze([
  '@deepseek-ai/dsh-attachment',
  '@deepseek-ai/dsh-subprocess',
  '@deepseek-ai/dsh-timeout',
  '@deepseek-ai/dsh-tools',
]);
// 工具链（运行时）的产地与版本：运行时不再是本仓库的产物，见 toolchain.lock.json。
const toolchainPin=JSON.parse(await readFile(join(root,'toolchain.lock.json'),'utf8'));
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
// 每次构建先清掉 lib/engines：它是上一版按当时的落点复制出来的，不清就会把旧落点的
// 脚本一起打进包里（实测过一次：pptx-office/engine/ 与新的 src/engine/ 同时存在，
// 旧的那份还带着过时的资源回推逻辑）。发布产物里只允许有当前布局。
await rm(join(coreOut,'lib','engines'),{recursive:true,force:true});
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
    const name=match[1],moduleRoot=resolve(dirname(args.path),'..');
    // 分区里引擎在 <模块根>/src/engine/、资源在 <模块根>/assets/，而引擎脚本用
    // Path(__file__).resolve().parents[2] 回推模块根找资源。所以 lib/engines/<name>/ 之下
    // 必须**照原样保留这个深度**（src/engine + assets），摊平会让 parents[2] 指到别处，
    // 表现是引擎在生成产物那一刻才 FileNotFoundError。文档模块没有相对资源，维持原样。
    const contents=(await readFile(args.path,'utf8'))
      .replaceAll('import.meta.url',`new URL('./engines/${name}/index.js', import.meta.url).href`)
      .replaceAll("'./engine/","'./src/engine/");
    await cp(join(moduleRoot,'src','engine'),join(coreOut,'lib','engines',name,'src','engine'),{recursive:true});
    if(await exists(join(moduleRoot,'assets')))await cp(join(moduleRoot,'assets'),join(coreOut,'lib','engines',name,'assets'),{recursive:true});
    return {contents,loader:'ts',resolveDir:dirname(args.path)};
  });
}};
// absWorkingDir 必须钉成 root：esbuild 的 metafile 输入键是相对**工作目录**的，默认取
// process.cwd()。从仓库目录之外调本脚本时，键会变成 "../开源团队作品/.../node_modules/..."
// 这种带 .. 的相对路径，下面那段 bundled 包统计用 resolve(root,key) 就一条都命中不到——
// 实测从 85 个包掉到 0，而 lib/bundled.json 一空，docx_doctor 的 exceljs 出处判断就失效、
// xlsx 退回假阴性。钉住之后键恒为仓库相对路径，与调用时的工作目录无关。
for(const entry of ['server','index'])await build({absWorkingDir:root,entryPoints:[join(root,'src',`${entry}.ts`)],outfile:join(coreOut,'lib',`${entry}.mjs`),bundle:true,platform:'node',format:'esm',target:'node22',external:['pdfjs-dist'],plugins:[engines],define:{__DSH_DOCX_VERSION__:JSON.stringify(VERSION)},banner:{js:"import { createRequire as __createRequire } from 'node:module'; const require = __createRequire(import.meta.url);"},metafile:true}).then(async result=>{
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
// 记下"哪些 npm 包被打进了 bundle"。doctor 用它来判断某个能力是否真的可用：
// exceljs 这类包是 inlined 的，磁盘上没有单独安装，从子进程 import 探测必然失败（假阴性）。
const bundledVersions={};
for(const [packageName,packageRoot] of bundledPackages){
  try{bundledVersions[packageName]=JSON.parse(await readFile(join(packageRoot,'package.json'),'utf8')).version??'unknown';}
  catch{bundledVersions[packageName]='unknown';}
}
await writeFile(join(coreOut,'lib','bundled.json'),JSON.stringify({schema:'dsh-office-bundled/v1',note:'这些 npm 包被 esbuild 打进了 lib/server.mjs，磁盘上没有单独安装；doctor 据此判断能力可用性。',packages:bundledVersions},null,2)+'\n');
console.log(`bundled packages recorded: ${Object.keys(bundledVersions).length}`);
await cp(join(root,'dsh'),join(coreOut,'dsh'),{recursive:true});
await cp(join(root,'dsh','cordis.patch.yml'),join(coreOut,'cordis.patch.yml'));
for(const name of ['ARCHITECTURE.md','OFFICE_ENGINE_DECISIONS.md','THIRD_PARTY.md','modules.lock.json'])await cp(join(root,name),join(coreOut,name));
// 核心包不再携带运行时，也不再限定平台：Windows 运行时是另一个包的事。
const coreManifest={name:'@deepseek-ai/dsh-docx',version:VERSION,type:'module',license:'UNLICENSED',description:'DOCX, PPTX, XLSX and native PDF modules with parsing, editing, rendering and versioned delivery for DSH. Runtime (Python/LibreOffice/Poppler) ships separately as @deepseek-ai/dsh-docx-runtime; without it the runtime-dependent capabilities report themselves unavailable instead of failing startup.',main:'dsh/index.mjs',exports:{'.':'./dsh/index.mjs','./core':'./lib/index.mjs','./package.json':'./package.json'},files:['dsh','lib','third-party-licenses','cordis.patch.yml','README.md','ARCHITECTURE.md','OFFICE_ENGINE_DECISIONS.md','THIRD_PARTY.md','modules.lock.json'],engines:{node:'^22.19.0 || >=24.0.0'},dsh:{bundle:{patch:'./cordis.patch.yml'}},dependencies:{
  '@deepseek-ai/dsh-mcp-client':DSH_MCP_VERSION,
  ...Object.fromEntries(DSH_MCP_STARTUP_DEPENDENCIES.map((name)=>[name,DSH_MCP_VERSION])),
  'pdfjs-dist':'6.3.289',
},peerDependencies:{'@deepseek-ai/cordis':'^4.0.1'}};
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
  toolchain:{repo:toolchainPin.repo,tag:toolchainPin.tag,asset:toolchainPin.asset,sha256:toolchainPin.sha256,note:'运行时的产地：按 dsh-office/toolchain.lock.json 钉住的工具链工件取出；本包只是把它按 <pkg>/runtime/<platform> 布局重新装好。'},
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
// 3. 元包：一步装好 = 核心 + 运行时。
//    依赖写**版本号**而不是 release 直链：pnpm 默认开 blockExoticSubdeps，子依赖里的 URL 依赖会被拒
//    （实测 ERR_PNPM_EXOTIC_SUBDEP），所以直链方案装不上。没有 registry 时这个元包装不起来——
//    发布形态因此是"分别装核心与运行时"，元包留着给"两个子包都发到 registry"的那天用。
// ---------------------------------------------------------------------------
await mkdir(fullOut,{recursive:true});
const fullPackage={name:'@deepseek-ai/dsh-docx-full',version:VERSION,type:'module',license:'UNLICENSED',description:'Meta package: the DSH office plugin core plus its Windows x64 runtime, for a one-step install.',private:false,dependencies:{
  '@deepseek-ai/dsh-docx':VERSION,
  '@deepseek-ai/dsh-docx-runtime':VERSION,
}};
await writeFile(join(fullOut,'package.json'),JSON.stringify(fullPackage,null,2)+'\n');
await writeFile(join(fullOut,'README.md'),`# @deepseek-ai/dsh-docx-full ${VERSION}

只写依赖的元包：\`@deepseek-ai/dsh-docx\`（核心）+ \`@deepseek-ai/dsh-docx-runtime\`（Windows x64 运行时）。

**它需要 registry**：依赖写的是版本号 ${VERSION}，两个子包目前只以 tarball 形式发布在本仓库的 Release 里，
没有被 registry 收录，所以从 tarball 装这个元包会解析不到依赖。试过把依赖改成 release 直链 URL，
但 pnpm 默认开 \`blockExoticSubdeps\`，会直接报 \`ERR_PNPM_EXOTIC_SUBDEP\`（子依赖不允许 URL 依赖）。

所以现在的装法是**分别装两个包**：

\`\`\`powershell
dsh plugin --profile web add '<核心>.tgz'
dsh plugin --profile web add '<运行时>.tgz'
\`\`\`

或者只装核心，用 \`runtimeRoot\` / \`DSH_OFFICE_RUNTIME_ROOT\` 指向你自己那份 LibreOffice/Python。
子包一旦发到 registry，这个元包就能直接当"一步装好"用。运行时的产地是
\`${toolchainPin.repo} ${toolchainPin.tag}\`（摘要钉在 dsh-office 的 \`toolchain.lock.json\`）。
`);
console.log(`full package built:    ${fullOut}`);

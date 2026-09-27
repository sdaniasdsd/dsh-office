import { access,realpath } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { dirname,join,resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const packageRoot=fileURLToPath(new URL('../',import.meta.url));
// 运行时包的名字：核心包只依赖它的**存在**，不把它列成硬依赖。
const runtimePackage='@deepseek-ai/dsh-docx-runtime';
const platformDir='win32-x64';
function windowsShortPath(path){
  // Do not send ordinary ASCII paths through cmd.exe: its /c quoting can turn
  // a drive-qualified path into a malformed path when the host drive differs.
  if(!/[\u0080-\uFFFF]/.test(path))return path;
  try{
    const output=execFileSync(process.env.ComSpec??'cmd.exe',['/d','/c','for %I in ("%THE_LAST_DOCX_SHORTPATH_INPUT%") do @echo %~sI'],{
      encoding:'utf8',windowsHide:true,env:{...process.env,THE_LAST_DOCX_SHORTPATH_INPUT:path},
    }).trim();
    return output&&!/[\u0080-\uFFFF]/.test(output)?output:path;
  }catch{return path;}
}
const exists=path=>access(path).then(()=>true,()=>false);
/**
 * 向上找兄弟运行时包。
 *
 * 为什么不用 import.meta.resolve：npm 的扁平布局与 pnpm 的隔离布局把包放在不同层级，
 * 而这个查找对两者都成立——从核心包真实路径一路向上，任何一层 node_modules 里有就认。
 */
async function findRuntimePackage(start){
  let dir=start;
  for(;;){
    const candidate=join(dir,'node_modules',...runtimePackage.split('/'));
    if(await exists(join(candidate,'runtime.json')))return candidate;
    const parent=dirname(dir);
    if(parent===dir)return null;
    dir=parent;
  }
}
/**
 * 运行时根目录的解析顺序：显式配置 → 环境变量 → 兄弟运行时包 → 随包自带的 runtime/。
 *
 * 这是"拆包"的接缝：核心包 14 MB，运行时是另一个包（或用户在别处装好的一份）。
 * 全都找不到时不抛错——缺的是**能力**，不是这个进程本身。
 */
async function resolveRuntimeRoot(config,root){
  const explicit=config.runtimeRoot??process.env.DSH_OFFICE_RUNTIME_ROOT;
  const source=config.runtimeRoot?'config':'env';
  if(explicit){
    const candidate=resolve(explicit);
    if(await exists(join(candidate,platformDir)))return {root:join(candidate,platformDir),source};
    return {root:candidate,source};
  }
  const sibling=await findRuntimePackage(root);
  if(sibling)return {root:join(sibling,'runtime',platformDir),source:'runtime-package',packageRoot:sibling};
  return {root:join(root,'runtime',platformDir),source:'bundled'};
}
/** Resolve once at activation; no ambient package downloads. */
export async function resolveDshConfig(config={},root=packageRoot){
  if(!config||typeof config!=='object'||Array.isArray(config))throw new Error('docx config must be an object');
  const allowed=new Set(['workspaceRoot','dataRoot','toolCallTimeoutMs','serverName','runtimeRoot','pythonPath','sofficePath','pdftoppmPath']);
  for(const key of Object.keys(config))if(!allowed.has(key))throw new Error(`Unknown docx configuration: ${key}`);
  for(const key of ['workspaceRoot','dataRoot','serverName'])if(config[key]!==undefined&&(typeof config[key]!=='string'||!config[key].trim()))throw new Error(`${key} must be a nonempty string`);
  for(const key of ['runtimeRoot','pythonPath','sofficePath','pdftoppmPath'])if(config[key]!==undefined&&(typeof config[key]!=='string'||!config[key].trim()))throw new Error(`${key} must be a nonempty string`);
  const workspace=await realpath(resolve(config.workspaceRoot??process.cwd()));
  const data=resolve(config.dataRoot??join(workspace,'.dsh-docx'));
  const timeout=config.toolCallTimeoutMs??180000;
  if(!Number.isInteger(timeout)||timeout<1000||timeout>600000)throw new Error('toolCallTimeoutMs must be 1000..600000');
  const serverName=config.serverName??'docx';
  if(!/^[A-Za-z0-9_-]{1,32}$/.test(serverName))throw new Error('Invalid serverName');
  const server=join(root,'lib','server.mjs');
  // 唯一致命的东西：这个包自己的服务端。缺它才是"包不完整"。
  if(!(await exists(server)))throw new Error(`DOCX package is incomplete: ${server}`);
  if(process.platform!=='win32'||process.arch!=='x64')console.warn(`[docx] 随包运行时只提供 Windows x64；当前 ${process.platform}/${process.arch}，需要运行时能力请用 runtimeRoot、DSH_OFFICE_RUNTIME_ROOT 或 @deepseek-ai/dsh-docx-runtime 指定。`);

  const runtime=await resolveRuntimeRoot(config,root);
  const python=config.pythonPath??process.env.DOCX_PYTHON??join(runtime.root,'python','python.exe');
  const soffice=config.sofficePath??process.env.DOCX_SOFFICE??windowsShortPath(join(runtime.root,'libreoffice','program','soffice.com'));
  const officeRuntime=join(runtime.root,'libreoffice','System64');
  const pdftoppm=config.pdftoppmPath??process.env.DOCX_PDFTOPPM??join(runtime.root,'poppler','poppler-26.09.0','Library','bin','pdftoppm.exe');
  const temporary=windowsShortPath(tmpdir());

  // 缺运行时**不再是启动错误**：逐个探明，只把存在的路径交给服务端，缺什么如实报出来。
  const binaries={python,soffice,pdftoppm};
  const found=Object.fromEntries(await Promise.all(Object.entries(binaries).map(async([name,path])=>[name,await exists(path)])));
  const missing=Object.entries(found).filter(([,present])=>!present).map(([name])=>`${name}=${binaries[name]}`);
  if(missing.length)console.warn(`[docx] 运行时不完整（来源 ${runtime.source}）：缺 ${missing.join('、')}。相关能力会在调用时明确报错，docx_doctor 会给出结论；其余能力不受影响。`);

  const env={THE_LAST_DOCX_WORKSPACE:workspace,THE_LAST_DOCX_DATA:data,TEMP:temporary,TMP:temporary,
    DSH_DOCX_RUNTIME_SOURCE:runtime.source,...(missing.length?{DSH_DOCX_RUNTIME_MISSING:missing.join(',')}:{})};
  if(found.python)env.DOCX_PYTHON=python;
  if(found.soffice)env.DOCX_SOFFICE=soffice;
  if(found.pdftoppm)env.DOCX_PDFTOPPM=pdftoppm;
  const pathEntries=[];
  if(await exists(officeRuntime))pathEntries.push(officeRuntime);
  if(process.env.PATH)pathEntries.push(process.env.PATH);
  if(pathEntries.length)env.PATH=pathEntries.join(';');
  return {transport:'stdio',serverName,command:process.execPath,args:[server],cwd:workspace,
    env,toolCallTimeoutMs:timeout,failOnStartupError:true,reconnect:{enabled:true,initialDelayMs:500,maxDelayMs:30000,maxAttempts:3}};
}

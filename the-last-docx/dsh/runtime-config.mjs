import { access,realpath } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { resolve,join } from 'node:path';
import { fileURLToPath } from 'node:url';
const packageRoot=fileURLToPath(new URL('../',import.meta.url));
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
/** Resolve once at activation; no ambient executable lookup or package downloads. */
export async function resolveDshConfig(config={},root=packageRoot){
  if(!config||typeof config!=='object'||Array.isArray(config))throw new Error('docx config must be an object');
  const allowed=new Set(['workspaceRoot','dataRoot','toolCallTimeoutMs','serverName']);
  for(const key of Object.keys(config))if(!allowed.has(key))throw new Error(`Unknown docx configuration: ${key}`);
  for(const key of ['workspaceRoot','dataRoot','serverName'])if(config[key]!==undefined&&(typeof config[key]!=='string'||!config[key].trim()))throw new Error(`${key} must be a nonempty string`);
  if(process.platform!=='win32'||process.arch!=='x64')throw new Error('This DOCX offline package supports Windows x64 only.');
  const workspace=await realpath(resolve(config.workspaceRoot??process.cwd()));
  const data=resolve(config.dataRoot??join(workspace,'.dsh-docx'));
  const timeout=config.toolCallTimeoutMs??180000;
  if(!Number.isInteger(timeout)||timeout<1000||timeout>600000)throw new Error('toolCallTimeoutMs must be 1000..600000');
  const serverName=config.serverName??'docx';
  if(!/^[A-Za-z0-9_-]{1,32}$/.test(serverName))throw new Error('Invalid serverName');
  const runtime=join(root,'runtime','win32-x64');
  const python=join(runtime,'python','python.exe');
  const soffice=windowsShortPath(join(runtime,'libreoffice','program','soffice.com'));
  const officeRuntime=join(runtime,'libreoffice','System64');
  const pdftoppm=join(runtime,'poppler','poppler-26.09.0','Library','bin','pdftoppm.exe');
  const server=join(root,'lib','server.mjs');
  const temporary=windowsShortPath(tmpdir());
  for(const path of [python,soffice,pdftoppm,server])await access(path).catch(()=>{throw new Error(`DOCX package is incomplete: ${path}`);});
  return {transport:'stdio',serverName,command:process.execPath,args:[server],cwd:workspace,
    env:{THE_LAST_DOCX_WORKSPACE:workspace,THE_LAST_DOCX_DATA:data,DOCX_PYTHON:python,DOCX_SOFFICE:soffice,DOCX_PDFTOPPM:pdftoppm,TEMP:temporary,TMP:temporary,PATH:`${officeRuntime}${process.env.PATH?`;${process.env.PATH}`:''}`},
    toolCallTimeoutMs:timeout,failOnStartupError:true,reconnect:{enabled:true,initialDelayMs:500,maxDelayMs:30000,maxAttempts:3}};
}

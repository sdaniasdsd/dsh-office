import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
/** 插件包根目录：跑 node 探测时用它做 cwd，让包自己的依赖能被解析到。 */
const packageRoot=fileURLToPath(new URL('../',import.meta.url));
/**
 * 被打进 bundle 的 npm 包（构建时由 scripts/build-dsh.mjs 写进 lib/bundled.json）。
 *
 * 为什么要这张表：exceljs 这类包是 inline 进 bundle 的，磁盘上没有单独安装，用子进程 import 去探测
 * 必然失败——那是**探测方法的假阴性**，不是能力缺失。从源码直接跑（没有这个文件）时表为空，
 * 行为退回"只看探测结果"，与改动前一致。
 */
function bundledPackages():Record<string,string>{
  try{return JSON.parse(readFileSync(fileURLToPath(new URL('../lib/bundled.json',import.meta.url)),'utf8')).packages??{};}
  catch{return {};}
}
export async function probe(command:string,args:string[],timeout=5000,cwd?:string):Promise<{available:boolean;detail:string}>{
  return new Promise(resolve=>{let settled=false;let output='';const child=spawn(command,args,{windowsHide:true,shell:false,stdio:['ignore','pipe','pipe'],...(cwd?{cwd}:{})});
    const finish=(available:boolean,detail:string)=>{if(settled)return;settled=true;clearTimeout(timer);resolve({available,detail});};
    const timer=setTimeout(()=>{child.kill();finish(false,'probe timeout');},timeout);
    child.stdout.on('data',v=>{output=(output+String(v)).slice(-2000);});child.stderr.on('data',v=>{output=(output+String(v)).slice(-2000);});
    child.once('error',()=>finish(false,'not found or cannot start'));child.once('close',code=>finish(code===0,output.trim()||`exit ${code}`));
  });
}
export async function doctor(config:{pythonPath?:string;sofficePath?:string;pdftoppmPath?:string}={}){
  const python=config.pythonPath??'python';
  const bundled=bundledPackages();
  // python 侧探测给 15 秒：这八条探测是并发起的，而 `import pptx`（连带 PIL/lxml）在冷启动
  // 加并发时会超过 5 秒——5 秒预算会把可用报成不可用（实测过一次 pptx:false 的假阴性）。
  const [runtime,xml,pptx,rdocx,docling,exceljsProbe,pdfjs,libreoffice,poppler]=await Promise.all([
    probe(python,['--version'],15000),probe(python,['-c','import lxml; print(lxml.__version__)'],15000),
    probe(python,['-c','import pptx; print(pptx.__version__)'],15000),
    probe(python,['-c','import importlib.util; print("rdocx found" if importlib.util.find_spec("rdocx") else "missing"); raise SystemExit(0 if importlib.util.find_spec("rdocx") else 1)'],15000),
    probe(python,['-c','import importlib.util; print("docling found" if importlib.util.find_spec("docling") else "missing"); raise SystemExit(0 if importlib.util.find_spec("docling") else 1)'],15000),
    // 这两条用插件包目录做 cwd：从服务端的工作目录解析，会看不到装在包里的依赖（假阴性）。
    probe(process.execPath,['-e','import("exceljs").then(m=>console.log(m.default?.Workbook ? "ExcelJS available" : "ExcelJS unavailable")).catch(()=>process.exit(1))'],5000,packageRoot),
    probe(process.execPath,['-e','import("pdfjs-dist/legacy/build/pdf.mjs").then(m=>console.log(m.version ? `PDF.js ${m.version}` : "PDF.js unavailable")).catch(()=>process.exit(1))'],5000,packageRoot),
    // soffice 首次启动（建用户 profile）常常超过 5 秒，5 秒预算会把可用报成不可用。
    probe(config.sofficePath??'soffice',['--version'],20000),probe(config.pdftoppmPath??'pdftoppm',['-v'],20000),
  ]);
  // exceljs 被 inline 进 bundle 时，探测失败不代表能力缺失；detail 里写明它的来路。
  const exceljsBundled=bundled['exceljs'];
  const exceljs={available:exceljsProbe.available||Boolean(exceljsBundled),
    detail:exceljsProbe.available?exceljsProbe.detail:`bundled into the server bundle${exceljsBundled?` (exceljs ${exceljsBundled})`:''}; not installed separately, so a child-process import cannot see it`};
  return {node:process.version,python:runtime,lxml:xml,pythonPptx:pptx,rdocx,docling,exceljs,pdfjs,libreoffice,poppler,
    bundledPackages:Object.keys(bundled),
    capabilities:{create:true,edit:true,artifact:runtime.available,parse:runtime.available,complexStructure:runtime.available,complexLayout:runtime.available&&rdocx.available,deepEasyParse:runtime.available&&docling.available,pptx:runtime.available&&pptx.available,xlsx:runtime.available&&exceljs.available,nativePdf:pdfjs.available,render:libreoffice.available&&poppler.available},
    note:'Availability is not an accuracy benchmark. Optional layout engines are never silently substituted. A bundled npm package is reported available with its provenance in "detail".'};
}

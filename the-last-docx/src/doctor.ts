import { spawn } from 'node:child_process';
export async function probe(command:string,args:string[],timeout=5000):Promise<{available:boolean;detail:string}>{
  return new Promise(resolve=>{let settled=false;let output='';const child=spawn(command,args,{windowsHide:true,shell:false,stdio:['ignore','pipe','pipe']});
    const finish=(available:boolean,detail:string)=>{if(settled)return;settled=true;clearTimeout(timer);resolve({available,detail});};
    const timer=setTimeout(()=>{child.kill();finish(false,'probe timeout');},timeout);
    child.stdout.on('data',v=>{output=(output+String(v)).slice(-2000);});child.stderr.on('data',v=>{output=(output+String(v)).slice(-2000);});
    child.once('error',()=>finish(false,'not found or cannot start'));child.once('close',code=>finish(code===0,output.trim()||`exit ${code}`));
  });
}
export async function doctor(config:{pythonPath?:string;sofficePath?:string;pdftoppmPath?:string}={}){
  const python=config.pythonPath??'python';
  const [runtime,xml,pptx,rdocx,docling,exceljs,libreoffice,poppler]=await Promise.all([
    probe(python,['--version']),probe(python,['-c','import lxml; print(lxml.__version__)']),
    probe(python,['-c','import pptx; print(pptx.__version__)']),
    probe(python,['-c','import importlib.util; print("rdocx found" if importlib.util.find_spec("rdocx") else "missing"); raise SystemExit(0 if importlib.util.find_spec("rdocx") else 1)']),
    probe(python,['-c','import importlib.util; print("docling found" if importlib.util.find_spec("docling") else "missing"); raise SystemExit(0 if importlib.util.find_spec("docling") else 1)']),
    probe(process.execPath,['-e','import("exceljs").then(m=>console.log(m.default?.Workbook ? "ExcelJS available" : "ExcelJS unavailable")).catch(()=>process.exit(1))']),
    probe(config.sofficePath??'soffice',['--version']),probe(config.pdftoppmPath??'pdftoppm',['-v']),
  ]);
  return {node:process.version,python:runtime,lxml:xml,pythonPptx:pptx,rdocx,docling,exceljs,libreoffice,poppler,
    capabilities:{create:true,edit:true,artifact:runtime.available,parse:runtime.available,complexStructure:runtime.available,complexLayout:runtime.available&&rdocx.available,deepEasyParse:runtime.available&&docling.available,pptx:runtime.available&&pptx.available,xlsx:runtime.available&&exceljs.available,render:libreoffice.available&&poppler.available},
    note:'Availability is not an accuracy benchmark. Optional layout engines are never silently substituted.'};
}

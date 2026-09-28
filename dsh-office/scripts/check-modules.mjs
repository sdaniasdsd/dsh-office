import { spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url));
// docx-runtime 是共享库（没有 module.json / 不注册 MCP 能力），但它是别的模块的编译依赖，
// 所以也要进这份清单：不然它的类型错误只会在别人 typecheck 时才暴露。
const modules=['docx-runtime','docx-inspect','docx-easy-parse','docx-parse','docx-complex-parse','docx-create','docx-styles','docx-edit','docx-render','docx-artifact','pptx-office','xlsx-office','pdf-office'];
const tsc=join(root,'node_modules/typescript/bin/tsc'),vitest=join(root,'node_modules/vitest/vitest.mjs');
let failed=false;
for(const name of modules){
  console.log(`Checking ${name}`);const cwd=join(root,'modules',name);
  for(const args of [[tsc,'--noEmit'],[vitest,'run']]){const r=spawnSync(process.execPath,args,{cwd,stdio:'inherit',windowsHide:true});if(r.status!==0)failed=true;}
}
process.exitCode=failed?1:0;

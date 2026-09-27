import { spawnSync } from 'node:child_process';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=fileURLToPath(new URL('..',import.meta.url));
const modules=['docx-inspect','docx-easy-parse','docx-parse','docx-complex-parse','docx-create','docx-styles','docx-edit','docx-render','docx-artifact','pptx-office','xlsx-office','pdf-office'];
const tsc=join(root,'node_modules/typescript/bin/tsc'),vitest=join(root,'node_modules/vitest/vitest.mjs');
let failed=false;
for(const name of modules){
  console.log(`Checking ${name}`);const cwd=join(root,'modules',name);
  for(const args of [[tsc,'--noEmit'],[vitest,'run']]){const r=spawnSync(process.execPath,args,{cwd,stdio:'inherit',windowsHide:true});if(r.status!==0)failed=true;}
}
process.exitCode=failed?1:0;

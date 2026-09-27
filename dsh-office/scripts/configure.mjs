import { writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve,join } from 'node:path';
const root=fileURLToPath(new URL('..',import.meta.url)),workspace=resolve(process.argv[2]??root);
const config={mcpServers:{'the-last-docx':{command:'node',args:[join(root,'scripts/start.mjs')],env:{THE_LAST_DOCX_WORKSPACE:workspace}}}};
await writeFile(join(root,'.mcp.json'),JSON.stringify(config,null,2)+'\n');
console.log(`Configured local plugin for workspace: ${workspace}. No app installation or global settings changed.`);

// Which copy of @deepseek-ai/dsh-mcp-client does the installed plugin actually get?
import { createRequire } from 'node:module';
const entry = 'C:\\Users\\AA\\AppData\\Roaming\\com.yeagoo.dsh-desktop\\harness\\profiles\\web\\node_modules\\@deepseek-ai\\dsh-docx\\dsh\\index.mjs';
const req = createRequire(entry);
console.log('resolved ->', req.resolve('@deepseek-ai/dsh-mcp-client'));

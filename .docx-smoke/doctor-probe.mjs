// ③-3 那组里 doctor 报 render:false 但真实渲染成功了：把它完整的探测结果打出来，看看是哪一项为假。
import { spawn } from 'node:child_process';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const ROOT = 'D:\\认真版agent\\.docx-smoke\\split-check';
const CORE = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx');
const { resolveDshConfig } = await import(pathToFileURL(join(CORE, 'dsh', 'runtime-config.mjs')).href);
const config = await resolveDshConfig({ workspaceRoot: 'C:\\Users\\AA\\Desktop\\bench' });
console.log(`runtime source: ${config.env.DSH_DOCX_RUNTIME_SOURCE}`);
console.log(`env: DOCX_PYTHON=${config.env.DOCX_PYTHON ? 'set' : 'unset'} DOCX_SOFFICE=${config.env.DOCX_SOFFICE ? 'set' : 'unset'} DOCX_PDFTOPPM=${config.env.DOCX_PDFTOPPM ? 'set' : 'unset'}`);

const child = spawn(process.execPath, [join(CORE, 'lib', 'server.mjs')], { env: { ...process.env, ...config.env }, stdio: ['pipe', 'pipe', 'pipe'] });
let buffer = ''; let id = 1; const pending = new Map();
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk;
  let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, index).trim(); buffer = buffer.slice(index + 1);
    if (!line) continue;
    let message; try { message = JSON.parse(line); } catch { continue; }
    if (message.id !== undefined && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  }
});
const request = (method, params) => { const current = id++; const promise = new Promise((resolve) => pending.set(current, resolve)); child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: current, method, params })}\n`); return promise; };
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'doctor-probe', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
const reply = await request('tools/call', { name: 'docx_doctor', arguments: {} });
const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
console.log('\ndocx_doctor 完整结果：');
console.log(JSON.stringify(JSON.parse(text), null, 1));
child.kill();
process.exit(0);

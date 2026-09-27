// 只回答一个问题：装好的 0.9.3 里，带版面几何的 docx_analyze 到底返回什么原始形状。
// 期望是「可诊断的稳定错误码」而不是崩溃的 ENGINE_FAILED。
import { spawn } from 'node:child_process';
import { join } from 'node:path';

const INSTALLED = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const SOURCE = join(WORKSPACE, 'sweep-093-source.docx');

const { resolveDshConfig } = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const config = await resolveDshConfig({ workspaceRoot: WORKSPACE });
const child = spawn(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { env: { ...process.env, ...config.env }, stdio: ['pipe', 'pipe', 'pipe'] });
let buffer = ''; let id = 1; const pending = new Map();
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk; let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const row = buffer.slice(0, index).trim(); buffer = buffer.slice(index + 1);
    if (!row) continue;
    let message; try { message = JSON.parse(row); } catch { continue; }
    if (message.id !== undefined && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  }
});
const request = (method, params) => { const current = id++; const promise = new Promise((resolve) => pending.set(current, resolve)); child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: current, method, params })}\n`); return promise; };
const callTool = async (name, args) => {
  const reply = await request('tools/call', { name, arguments: args });
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  return reply.error ? { transportError: reply.error } : { isError: reply.result?.isError === true, raw: text.slice(0, 700) };
};
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'probe-layout', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
const imported = await callTool('docx_import', { path: SOURCE });
const refId = imported.raw.match(/"id":"(sha256:[0-9a-f]+)"/)?.[1];
console.log(`ref=${refId}`);
console.log(`layout   -> ${JSON.stringify(await callTool('docx_analyze', { artifactRef: { id: refId, uri: 'ignored' }, requestId: 'probe-layout', complex: 'layout' }))}`);
console.log(`structure-> ${(await callTool('docx_analyze', { artifactRef: { id: refId, uri: 'ignored' }, requestId: 'probe-structure', complex: 'structure' })).raw.slice(0, 200)}`);
child.kill();
process.exit(0);

// What can the installed pptx-office module actually do? Ask the plugin's own
// registry instead of guessing from a tool description.
import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const PACKAGE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
const child = spawn(process.execPath, [join(PACKAGE, 'lib', 'server.mjs')], {
  env: { ...process.env, THE_LAST_DOCX_WORKSPACE: 'C:\\Users\\AA\\Desktop\\bench', THE_LAST_DOCX_DATA: 'D:\\认真版agent\\.docx-smoke\\data-pptx' },
  stdio: ['pipe', 'pipe', 'pipe'],
});
let nextId = 1;
const pending = new Map();
let buffer = '';
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk;
  let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, index).trim();
    buffer = buffer.slice(index + 1);
    if (!line) continue;
    let message;
    try { message = JSON.parse(line); } catch { continue; }
    if (message.id !== undefined && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  }
});
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => process.stderr.write(chunk));
const request = (method, params) => {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
};
const notify = (method, params) => child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'pptx-capabilities', version: '1' } });
notify('notifications/initialized');
const reply = await request('tools/call', { name: 'docx_modules', arguments: {} });
const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
const registry = JSON.parse(text);
const pptx = registry.find((entry) => entry.registryId === 'pptx-office');
console.log(`modules in registry: ${registry.map((entry) => entry.registryId).join(', ')}\n`);
console.log(JSON.stringify(pptx, null, 2));
child.kill();
process.exit(0);

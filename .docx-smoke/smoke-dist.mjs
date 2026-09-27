// 构建产物的冒烟测试：拉起 dist 里的 server，握手 + 列模块注册表，确认四个家族都在。
import { spawn } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const PACKAGE = process.argv[2];
const child = spawn(process.execPath, [join(PACKAGE, 'lib', 'server.mjs')], {
  env: { ...process.env, THE_LAST_DOCX_WORKSPACE: 'D:\\认真版agent\\.docx-smoke\\data-pptx', THE_LAST_DOCX_DATA: 'D:\\认真版agent\\.docx-smoke\\data-smoke' },
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
child.stderr.on('data', (chunk) => process.stderr.write(`[server] ${chunk}`));
const request = (method, params) => {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
};
const notify = (method, params) => child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);

console.log(`package ${PACKAGE}`);
console.log(`  version ${JSON.parse(readFileSync(join(PACKAGE, 'package.json'), 'utf8')).version}`);
const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'smoke', version: '1' } });
notify('notifications/initialized');
console.log(`  handshake ${JSON.stringify(initialize.result?.serverInfo)}`);
const listed = await request('tools/list', {});
console.log(`  tools (${(listed.result?.tools ?? []).length}): ${(listed.result?.tools ?? []).map((tool) => tool.name).join(', ')}`);
const modules = await request('tools/call', { name: 'docx_modules', arguments: {} });
const registry = JSON.parse(modules.result?.content?.find((part) => part.type === 'text')?.text ?? '[]');
console.log(`  modules (${registry.length}): ${registry.map((entry) => entry.registryId).join(', ')}`);
const byGroup = {};
for (const entry of registry) byGroup[entry.definition?.profileGroup ?? '?'] = (byGroup[entry.definition?.profileGroup ?? '?'] ?? 0) + 1;
console.log(`  by family: ${JSON.stringify(byGroup)}`);
child.kill();
process.exit(registry.length === 12 ? 0 : 1);

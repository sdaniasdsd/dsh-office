// 回答「能不能只靠 Node 生态」需要两个事实：包里到底带了哪些非 Node 运行时（各占多大），
// 以及每个模块在注册表里声明的运行时依赖是什么（哪些模块本来就只用 Node）。
import { spawn } from 'node:child_process';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const PACKAGE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');

const size = (path) => {
  let bytes = 0; let files = 0;
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const child = join(dir, entry.name);
      if (entry.isDirectory()) walk(child);
      else if (entry.isFile()) { bytes += statSync(child).size; files += 1; }
    }
  };
  walk(path);
  return { bytes, files };
};

console.log('=== 随包分发的非 Node 运行时（安装后的 0.8.1） ===');
for (const name of ['python', 'libreoffice', 'poppler']) {
  const path = join(PACKAGE, 'runtime', 'win32-x64', name);
  try {
    const { bytes, files } = size(path);
    console.log(`  ${name.padEnd(12)} ${String(Math.round(bytes / 1048576)).padStart(6)} MB  ${String(files).padStart(6)} files`);
  } catch { console.log(`  ${name.padEnd(12)} (missing)`); }
}
const lib = size(join(PACKAGE, 'lib'));
console.log(`  ${'lib (JS)'.padEnd(12)} ${String(Math.round(lib.bytes / 1048576)).padStart(6)} MB  ${String(lib.files).padStart(6)} files`);
const total = size(PACKAGE);
console.log(`  ${'包总计'.padEnd(12)} ${String(Math.round(total.bytes / 1048576)).padStart(6)} MB  ${String(total.files).padStart(6)} files`);

// 注册表里的运行时依赖：哪些模块需要 python / LibreOffice / poppler
const child = spawn(process.execPath, [join(PACKAGE, 'lib', 'server.mjs')], {
  env: { ...process.env, THE_LAST_DOCX_WORKSPACE: 'D:\\认真版agent\\.docx-smoke\\data-pptx', THE_LAST_DOCX_DATA: 'D:\\认真版agent\\.docx-smoke\\data-deps' },
  stdio: ['pipe', 'pipe', 'pipe'],
});
let buffer = '';
let nextId = 1;
const pending = new Map();
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
child.stderr.on('data', () => {});
const request = (method, params) => {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
};
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'node-only-audit', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
const reply = await request('tools/call', { name: 'docx_modules', arguments: {} });
const registry = JSON.parse(reply.result.content.find((part) => part.type === 'text').text);

console.log('\n=== 每个模块声明的运行时依赖 ===');
const classify = new Map();
for (const entry of registry) {
  const deps = entry.definition.dependencies ?? [];
  const runtimes = deps.filter((dep) => dep.kind === 'runtime').map((dep) => `${dep.name}${dep.optional ? '(optional)' : ''}`);
  const npm = deps.filter((dep) => dep.kind === 'npm').map((dep) => `${dep.name}@${dep.version ?? ''}`);
  console.log(`  ${entry.registryId.padEnd(20)} runtimes: ${runtimes.length ? runtimes.join(', ') : '（无，纯 Node/TS）'}${npm.length ? `   npm: ${npm.join(', ')}` : ''}`);
  const key = runtimes.length ? runtimes.join('+') : '(none)';
  classify.set(key, (classify.get(key) ?? 0) + 1);
}
console.log('\n=== 归类 ===');
for (const [key, count] of [...classify].sort((a, b) => b[1] - a[1])) console.log(`  ${String(count).padStart(2)} 个模块：${key}`);
child.kill();
process.exit(0);

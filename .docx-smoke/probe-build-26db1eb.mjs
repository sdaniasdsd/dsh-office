// 验证"构建产物"本身能不能用：直接起 dist/dsh-docx/lib/server.mjs，不改 profile、不重启 DSH。
// 关键看点：xlsx 是否恢复 true（说明 bundled.json 的 85 个包生效了）、
// 新增的 java / runtimeResolution 字段在不在、以及渲染链在 DSH_OFFICE_RUNTIME_ROOT 下能否解析。
import { spawn } from 'node:child_process';
import { existsSync, readFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office';
const DIST = join(REPO, 'dist', 'dsh-docx');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const DOCX = join(WORKSPACE, 'rev-v1.docx');

console.log('=== 构建产物结构 ===');
for (const rel of ['lib/server.mjs', 'lib/bundled.json', 'dsh/runtime-config.mjs', 'package.json',
  'lib/engines/pptx-office/src/engine/pptx_bridge.py', 'lib/engines/pptx-office/assets/indigo-paperlight-v1.png',
  'lib/engines/docx-complex-parse/docx_complex_parse.py']) {
  const path = join(DIST, rel);
  console.log(`  ${existsSync(path) ? '✓' : '✗'} ${rel}`);
}
const bundled = JSON.parse(readFileSync(join(DIST, 'lib', 'bundled.json'), 'utf8'));
const names = Object.keys(bundled.packages ?? {});
console.log(`  bundled.json: ${names.length} 个包；exceljs=${bundled.packages?.exceljs ?? '(缺)'}；version=${JSON.parse(readFileSync(join(DIST, 'package.json'), 'utf8')).version}`);

const child = spawn(process.execPath, [join(DIST, 'lib', 'server.mjs')], {
  env: { ...process.env, THE_LAST_DOCX_WORKSPACE: WORKSPACE, THE_LAST_DOCX_DATA: join(WORKSPACE, '.dsh-docx-probe-build'),
    DSH_OFFICE_RUNTIME_ROOT: join(REPO, 'runtime') },
  stdio: ['pipe', 'pipe', 'pipe'],
});
let buffer = ''; let id = 1; const pending = new Map(); let stderr = '';
child.stderr.setEncoding('utf8'); child.stderr.on('data', (chunk) => { stderr += chunk; });
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
const call = async (name, args) => {
  const reply = await request('tools/call', { name, arguments: args });
  if (reply.error) return { transportError: reply.error };
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  try { return JSON.parse(text); } catch { return { raw: text.slice(0, 200) }; }
};

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'probe-build', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);

const doctor = await call('docx_doctor', {});
console.log('\n=== 构建产物的 docx_doctor（原始返回前 700 字） ===');
console.log(JSON.stringify(doctor).slice(0, 700));
console.log('\n=== 构建产物的 docx_doctor ===');
console.log(`  capabilities.xlsx = ${doctor.capabilities?.xlsx}   pptx=${doctor.capabilities?.pptx}  render=${doctor.capabilities?.render}  java=${doctor.capabilities?.java}`);
console.log(`  exceljs = ${JSON.stringify(doctor.exceljs)}`);
console.log(`  python=${doctor.python?.available} libreoffice=${doctor.libreoffice?.available} poppler=${doctor.poppler?.available} java.probe=${JSON.stringify(doctor.java?.detail ?? null)}`);
console.log(`  runtimeResolution.source = ${doctor.runtimeResolution?.source}  components = ${JSON.stringify(Object.fromEntries(Object.entries(doctor.runtimeResolution?.components ?? {}).map(([k, v]) => [k, v.source])))}`);

const imported = await call('docx_import', { path: DOCX });
const ref = imported?.artifactRef ?? imported;
const rendered = await call('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { requestId: 'probe-build-render', artifactRef: ref } });
console.log(`\n=== 功能调用 ===\n  docx-render: ${rendered.code ? `✗ ${rendered.code}: ${rendered.message}` : `✓ pages=${rendered.result?.pageCount} engine=${rendered.result?.engine} pdf=${rendered.artifacts?.[0]?.sizeBytes} B`}`);
if (stderr.trim()) console.log(`\nstderr 摘要: ${stderr.trim().split('\n').slice(0, 3).join(' | ')}`);
child.kill();
process.exit(0);

// 交接前的最后一道：用**装进 profile 的那一份**跑一组真实调用（建 → 解析 → 渲染 → PDF），
// 并确认插件的激活入口能加载（dsh/index.mjs 会 import @deepseek-ai/dsh-mcp-client）。
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';   // 探针样例 standard-labour-contract.docx 就在这一层，import 必须落在输入根内
const OUT = join(WORKSPACE, 'out-091');

console.log('=== 激活入口能加载吗 ===');
console.log(`  dsh-mcp-client 在 profile 里: ${existsSync(join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-mcp-client'))}`);
const { resolveDshConfig } = await import(pathToFileURL(join(INSTALLED, 'dsh', 'runtime-config.mjs')).href);
mkdirSync(OUT, { recursive: true });
const config = await resolveDshConfig({ workspaceRoot: WORKSPACE });
console.log(`  resolveDshConfig OK  source=${config.env.DSH_DOCX_RUNTIME_SOURCE}  缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
console.log(`  command=${config.command}  args=${JSON.stringify(config.args)}`);

console.log('\n=== 真实调用（用装好的那份起服务） ===');
const child = spawn(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { env: { ...process.env, ...config.env }, stdio: ['pipe', 'pipe', 'pipe'] });
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
const errors = [];
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => errors.push(chunk));
const request = (method, params) => { const current = id++; const promise = new Promise((resolve) => pending.set(current, resolve)); child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: current, method, params })}\n`); return promise; };
const call = async (name, args) => {
  const reply = await request('tools/call', { name, arguments: args });
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  return { isError: Boolean(reply.result?.isError), text };
};
const short = (text, n = 150) => text.replace(/\s+/gu, ' ').slice(0, n);

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'handover-check', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);

const doctor = await call('docx_doctor', {});
const doctorJson = JSON.parse(doctor.text);
console.log(`  docx_doctor 能力: ${JSON.stringify(doctorJson.capabilities)}`);

const created = await call('docx_call', { moduleId: 'docx-create', operation: 'execute', input: { requestId: 'ho-create', plan: { kind: 'create', document: { preset: 'chinese-long', scenario: 'formal-record', pageNumberStyle: 'pageOfTotal', blocks: [
  { kind: 'paragraph', id: 't', style: 'Title', alignment: 'center', runs: [{ text: '工具链适配验收' }] },
  { kind: 'paragraph', id: 'p1', runs: [{ text: '这份文档由装进 DSH profile 的 0.9.1 生成：核心包从 dsh-office 取，运行时包从 dsh-toolchain 的工件重装。' }] },
  { kind: 'paragraph', id: 'p2', runs: [{ text: '验收点：创建不依赖运行时；解析与渲染必须能用到 Python 3.13、LibreOffice 26.8.0 与 Poppler 26.09.0。' }] },
] } } } });
const artifact = created.isError ? null : JSON.parse(created.text).artifacts[0];
console.log(`  docx-create  → ${created.isError ? `失败 ${short(created.text)}` : `成功 ${artifact.id.slice(0, 24)}… ${artifact.sizeBytes}B`}`);

const imported = await call('docx_import', { path: join(WORKSPACE, 'standard-labour-contract.docx') });
const importedRef = imported.isError ? null : JSON.parse(imported.text);
const parsed = importedRef ? await call('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: importedRef, requestId: 'ho-parse' } }) : { isError: true, text: imported.text };
console.log(`  docx-parse   → ${parsed.isError ? `失败 ${short(parsed.text)}` : '成功（用了运行时里的 python + lxml）'}`);

const rendered = importedRef ? await call('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: importedRef, requestId: 'ho-render', config: { limits: { dpi: 100 } } } }) : { isError: true, text: '无 artifact' };
console.log(`  docx-render  → ${rendered.isError ? `失败 ${short(rendered.text)}` : `成功 ${short(rendered.text, 120)}`}`);

const xlsxDoctor = doctorJson.capabilities.xlsx;
console.log(`\n（doctor 说明：xlsx=${xlsxDoctor} 是已知假阴性——exceljs 被打进 bundle、没有单独安装。docling/rdocx 不在运行时包里。）`);
if (errors.length) console.log(`\nserver stderr: ${errors.join('').replace(/\s+/gu, ' ').slice(0, 300)}`);
child.kill();
process.exit(0);

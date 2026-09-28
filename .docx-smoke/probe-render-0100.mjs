// 渲染链单独复测：换掉带外部超链接的样本（它会撞 render 的 allowExternalLinks:false 默认策略），
// 确认新核心在新 resolver 下 soffice/pdftoppm 路径确实解析到并产出 PDF。
import { spawn } from 'node:child_process';
import { join } from 'node:path';
import { existsSync } from 'node:fs';

const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';

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
const call = async (name, args) => {
  const reply = await request('tools/call', { name, arguments: args });
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  try { return JSON.parse(text); } catch { return { raw: text.slice(0, 300) }; }
};
const outcome = (r) => (r?.code ? `✗ ${r.code}: ${r.message ?? ''}` : null);
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'probe-render', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);

for (const name of ['case-02-project-status.docx', 'rev-v1.docx']) {
  const source = join(WORKSPACE, name);
  if (!existsSync(source)) { console.log(`${name}: 不在工作区，跳过`); continue; }
  const imported = await call('docx_import', { path: source });
  const ref = imported?.artifactRef ?? imported;
  if (!ref?.id) { console.log(`${name}: 导入失败 ${JSON.stringify(imported).slice(0, 160)}`); continue; }
  const rendered = await call('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { requestId: `render-${name}`, artifactRef: ref } });
  console.log(`${name}: ${outcome(rendered) ?? `✓ pages=${rendered.result?.pageCount} engine=${rendered.result?.engine} 产物=${JSON.stringify((rendered.artifacts ?? []).map((a) => `${a.label}:${a.sizeBytes}B`))}`}`);
}
child.kill();
process.exit(0);

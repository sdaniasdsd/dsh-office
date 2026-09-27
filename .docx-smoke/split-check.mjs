// 拆包验证：核心包当作用户装好的样子放进去，看两件事——
//   ① 有兄弟运行时包时，核心包能不能自己找到并拿到三条路径；
//   ② 没有运行时包时，核心包**不炸**、如实报缺，而且不依赖运行时的能力仍然可用。
//
// 用真实拷贝（核心只有 14 MB）而不是软链接：import.meta.url 走真实路径，
// 软链接会把"向上找 node_modules"那段逻辑绕过去，测不到真实行为。
import { spawn } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, rmSync, symlinkSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const REPO = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office';
const DIST = join(REPO, 'dist');
const ROOT = 'D:\\认真版agent\\.docx-smoke\\split-check';
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const SAMPLE = join(WORKSPACE, 'standard-labour-contract.docx');
const CORE = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx');
const RUNTIME_PKG = join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-docx-runtime');

rmSync(ROOT, { recursive: true, force: true });
mkdirSync(join(ROOT, 'node_modules', '@deepseek-ai'), { recursive: true });
cpSync(join(DIST, 'dsh-docx'), CORE, { recursive: true });
// 核心包的两条外部依赖：真实项目里由 npm 装，这里指回仓库里已装好的那份。
symlinkSync(join(REPO, 'node_modules', 'pdfjs-dist'), join(ROOT, 'node_modules', 'pdfjs-dist'), 'junction');
symlinkSync(join(REPO, 'node_modules', '@deepseek-ai', 'dsh-mcp-client'), join(ROOT, 'node_modules', '@deepseek-ai', 'dsh-mcp-client'), 'junction');

const { resolveDshConfig } = await import(pathToFileURL(join(CORE, 'dsh', 'runtime-config.mjs')).href);

const describe = (config, label) => {
  console.log(`\n=== ${label} ===`);
  console.log(`  runtime source = ${config.env.DSH_DOCX_RUNTIME_SOURCE}`);
  console.log(`  缺失项         = ${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
  for (const key of ['DOCX_PYTHON', 'DOCX_SOFFICE', 'DOCX_PDFTOPPM']) {
    const value = config.env[key];
    console.log(`  ${key.padEnd(14)} = ${value && existsSync(value) ? '存在' : '(未注入或不存在)'}`);
  }
};

// ---- ① 只有核心包：不炸，明确报缺 -------------------------------------------
let degraded;
try {
  degraded = await resolveDshConfig({ workspaceRoot: WORKSPACE });
  describe(degraded, '① 只有核心包（无兄弟运行时包）');
} catch (error) {
  console.log(`\n=== ① 只有核心包 ===\n  ✗ 抛错了：${error.message}`);
  process.exit(1);
}

// ---- ② 装上运行时包：核心包自己找到 -----------------------------------------
symlinkSync(join(DIST, 'dsh-docx-runtime'), RUNTIME_PKG, 'junction');
const withRuntime = await resolveDshConfig({ workspaceRoot: WORKSPACE });
describe(withRuntime, '② 核心 + 兄弟运行时包');

// ---- ③ 两种配置下真实调用 ----------------------------------------------------
function connect(serverEnv) {
  const child = spawn(process.execPath, [join(CORE, 'lib', 'server.mjs')], { env: { ...process.env, ...serverEnv }, stdio: ['pipe', 'pipe', 'pipe'] });
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
  return { child, errors, request };
}
const shorten = (text, limit = 190) => text.replace(/\s+/gu, ' ').slice(0, limit);

for (const [label, serverEnv] of [
  ['③-1 只有核心包（保留宿主 PATH）', degraded.env],
  ['③-2 只有核心包（PATH 收成 System32，模拟干净机器）', { ...degraded.env, PATH: 'C:\\Windows\\System32' }],
  ['③-3 核心 + 运行时包', withRuntime.env],
]) {
  console.log(`\n=== ${label}：真实调用 ===`);
  const handle = connect(serverEnv);
  await handle.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'split-check', version: '1' } });
  handle.child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
  const call = async (name, args) => {
    const reply = await handle.request('tools/call', { name, arguments: args });
    const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
    return { isError: Boolean(reply.result?.isError), text };
  };
  const grade = (result) => (result.isError ? '失败' : '成功');

  const doctor = await call('docx_doctor', {});
  console.log(`  docx_doctor                   → ${grade(doctor)}  ${shorten(doctor.text, 260)}`);

  const created = await call('docx_call', { moduleId: 'docx-create', operation: 'execute', input: { requestId: 'split-create', plan: { kind: 'create', document: { preset: 'chinese-long', blocks: [{ kind: 'paragraph', id: 'p1', runs: [{ text: '拆包验证：创建文档不需要运行时。' }] }] } } } });
  console.log(`  docx-create      (无运行时)    → ${grade(created)}  ${shorten(created.text, 110)}`);

  const imported = await call('docx_import', { path: SAMPLE });
  const artifact = imported.isError ? null : JSON.parse(imported.text);
  const parsed = artifact
    ? await call('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: artifact, requestId: 'split-parse' } })
    : { isError: true, text: `import failed: ${shorten(imported.text, 100)}` };
  console.log(`  docx-parse execute (python)    → ${grade(parsed)}  ${shorten(parsed.text, 170)}`);

  const importedPptx = await call('docx_import', { path: join(WORKSPACE, 'out-speech', '嵌入式的发展路径-配套PPT.pptx') });
  const pptxArtifact = importedPptx.isError ? null : JSON.parse(importedPptx.text);
  const pptx = pptxArtifact
    ? await call('pptx_call', { operation: 'inspect', input: { requestId: 'split-pptx', operation: 'inspect', artifactRef: pptxArtifact } })
    : { isError: true, text: `import failed: ${shorten(importedPptx.text, 80)}` };
  console.log(`  pptx inspect   (python-pptx)   → ${grade(pptx)}  ${shorten(pptx.text, 150)}`);

  const rendered = artifact
    ? await call('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: artifact, requestId: 'split-render', config: { limits: { dpi: 72 } } } })
    : { isError: true, text: 'no artifact' };
  console.log(`  docx-render   (LibreOffice)    → ${grade(rendered)}  ${shorten(rendered.text, 170)}`);

  handle.child.kill();
}
console.log('\n结论看两条线：无运行时那组里「无运行时」标记的能力应当失败且报清楚，其余应当成功；有运行时那组应全部成功。');

// 取真实回执（修正后的工具形状），并验证分区新模块 docx-runtime 的解析能力：
//   A) 按 DSH 的方式注入环境
//   B) 剥掉 DOCX_* / DSH_OFFICE_RUNTIME_ROOT：外壳仍会注入裸名 'python'/'soffice'，
//      但模块 resolver 对直接配置会做 existsSync 判定，裸名判不存在 → 应回落到兄弟包。
import { spawn } from 'node:child_process';
import { join } from 'node:path';
import { copyFileSync, mkdirSync } from 'node:fs';

const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const DOCX_SOURCE = join(WORKSPACE, 'upstream-0100-source.docx');
const PPTX_SOURCE = join(WORKSPACE, 'out-speech', '嵌入式的发展路径-配套PPT.pptx');

const { resolveDshConfig } = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const shellConfig = await resolveDshConfig({ workspaceRoot: WORKSPACE });

const start = (env) => {
  const child = spawn(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { env, stdio: ['pipe', 'pipe', 'pipe'] });
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
  return { child, request };
};
const callTool = async (session, name, args) => {
  const reply = await session.request('tools/call', { name, arguments: args });
  if (reply.error) return { transportError: reply.error };
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  try { return JSON.parse(text); } catch { return { raw: text.slice(0, 300) }; }
};
const handshake = async (session, label) => {
  await session.request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: label, version: '1' } });
  session.child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
};
const importId = async (session, path) => {
  const result = await callTool(session, 'docx_import', { path });
  return JSON.stringify(result).match(/"id":"(sha256:[0-9a-f]+)"/)?.[1] ?? `✗ ${JSON.stringify(result).slice(0, 200)}`;
};
// 引用必须逐字用导入回执本身：只给 {id, uri:'占位'} 会被服务端拒为 ARTIFACT_UNAVAILABLE，
// 而那种失败长得像成功——所以这里同时返回 ref 对象与一句人读的状态。
const importRef = async (session, path) => {
  const result = await callTool(session, 'docx_import', { path });
  const ref = result?.artifactRef ?? (result && typeof result === 'object' && typeof result.id === 'string' && typeof result.uri === 'string' ? result : null);
  return { ref, note: ref ? `${ref.id} (${ref.sizeBytes ?? '?'} B)` : `✗ ${JSON.stringify(result).slice(0, 240)}` };
};
// 模块错误以 {code, message} 的形状回来，不是 {error}。别把它当成功。
const outcome = (result) => (result?.code ? `✗ ${result.code}: ${result.message ?? ''}` : null);

mkdirSync(WORKSPACE, { recursive: true });
const pptxTarget = join(WORKSPACE, 'upstream-0100-deck.pptx');
copyFileSync(PPTX_SOURCE, pptxTarget);

console.log('=== A：外壳注入环境（DSH 实际用法） ===');
const a = start({ ...process.env, ...shellConfig.env });
await handshake(a, 'probe-A');

const pptx = await importRef(a, pptxTarget);
console.log(`pptx 导入: ${pptx.note}`);
const styled = await callTool(a, 'pptx_call', { operation: 'execute', input: { requestId: 'probe-a-art', artifactRef: pptx.ref, payload: { action: 'applyArtStyle', styleId: 'indigo-paperlight-v1', artWord: 'DSH' } } });
console.log(`applyArtStyle: ${outcome(styled) ?? JSON.stringify({ styleId: styled.result?.styleId, styledSlides: styled.result?.styledSlides, artWord: styled.result?.artWord, slideCount: styled.result?.slideCount })}`);
console.log(`  产物: ${JSON.stringify((styled.artifacts ?? []).map((x) => ({ id: x.id, bytes: x.sizeBytes, label: x.label })))}`);
const styledId = styled.artifacts?.[0];
if (styledId) {
  const verifyStyled = await callTool(a, 'pptx_call', { operation: 'verify', input: { requestId: 'probe-a-art-verify', artifactRef: styledId, payload: { expectedSlideCount: 12 } } });
  console.log(`  改写后 verify: ${outcome(verifyStyled) ?? JSON.stringify(verifyStyled.result?.summary ?? verifyStyled.result)}`);
}

const docx = await importRef(a, DOCX_SOURCE);
const rendered = await callTool(a, 'docx_call', { moduleId: 'docx-render', operation: 'execute', input: { requestId: 'probe-a-render', artifactRef: docx.ref } });
console.log(`docx-render: ${outcome(rendered) ?? JSON.stringify({ pageCount: rendered.result?.pageCount, engine: rendered.result?.engine, pdfBytes: rendered.artifacts?.[0]?.sizeBytes })}`);
const analyzed = await callTool(a, 'docx_analyze', { artifactRef: docx.ref, requestId: 'probe-a-analyze', complex: 'structure' });
console.log(`docx_analyze: ${outcome(analyzed) ?? `✓ bridges=${analyzed.bridges?.length} complexMode=${analyzed.complexMode}`}`);
a.child.kill();

console.log('\n=== B：剥掉 DOCX_* / DSH_OFFICE_RUNTIME_ROOT（只留最小环境） ===');
const minimal = { SystemRoot: process.env.SystemRoot, windir: process.env.windir, TEMP: process.env.TEMP, TMP: process.env.TMP, PATH: 'C:\\Windows\\System32', THE_LAST_DOCX_WORKSPACE: WORKSPACE, THE_LAST_DOCX_DATA: join(WORKSPACE, '.dsh-docx-probe-b') };
const b = start(minimal);
await handshake(b, 'probe-B');
const docxB = await importRef(b, DOCX_SOURCE);
const renderB = await callTool(b, 'docx_call', { moduleId: 'docx-render', operation: 'execute', input: { requestId: 'probe-b-render', artifactRef: docxB.ref } });
console.log(`docx-render: ${outcome(renderB) ?? `✓ pages=${renderB.result?.pageCount} engine=${renderB.result?.engine} pdf=${renderB.artifacts?.[0]?.sizeBytes} B`}`);
const parsedB = await callTool(b, 'docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { requestId: 'probe-b-parse', artifactRef: docxB.ref } });
console.log(`docx-parse : ${outcome(parsedB) ?? `✓ blocks=${parsedB.result?.ir?.blocks?.length} engine=${parsedB.telemetry?.engine}`}`);
b.child.kill();
process.exit(0);

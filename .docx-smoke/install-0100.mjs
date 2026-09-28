// 装最新插件核心包（0.10.0：装配了分区新模块 docx-runtime + pptx 封面样式），
// 运行时包沿用已装好的那份（内容与适配器 v0.9.0 资产一致），然后回读：
//   版本 / 握手 / 运行时解析 / doctor 能力 / 新 resolver 在 docx-render 上的实际路径解析 /
//   pptx 新能力 applyArtStyle（这条同时验证打包布局把资源放对了地方）。
import { spawn, spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const CORE_TGZ = join(DIST, 'deepseek-ai-dsh-docx-0.10.0.tgz');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const DOCX_SOURCE = join(WORKSPACE, 'sweep-093-source.docx');
const PPTX_SOURCE = join(WORKSPACE, 'out-speech', '嵌入式的发展路径-配套PPT.pptx');

if (!existsSync(CORE_TGZ)) throw new Error(`缺 tarball：${CORE_TGZ}`);
copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-0.10.0`);
const profile = JSON.parse(readFileSync(PROFILE_PACKAGE, 'utf8'));
const previous = profile.dependencies['@deepseek-ai/dsh-docx'];
profile.dependencies['@deepseek-ai/dsh-docx'] = `file:${CORE_TGZ.replaceAll('\\', '/')}`;
writeFileSync(PROFILE_PACKAGE, `${JSON.stringify(profile, null, 4)}\n`);
console.log(`核心依赖：was ${previous}\n         now ${profile.dependencies['@deepseek-ai/dsh-docx']}`);
console.log(`运行时依赖保持：${profile.dependencies['@deepseek-ai/dsh-docx-runtime']}`);

for (let attempt = 1; attempt <= 2; attempt += 1) {
  console.log(`\n$ pnpm install --prefer-offline  (第 ${attempt} 次)`);
  const result = spawnSync('pnpm', ['install', '--prefer-offline'], { cwd: PROFILE, stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
  if (result.status === 0) break;
  if (attempt === 2) throw new Error('pnpm install 两次都没成');
}

console.log(`\n核心版本: ${JSON.parse(readFileSync(join(INSTALLED, 'package.json'), 'utf8')).version}`);
// 打包布局回执：分区要求 <模块根>/assets，引擎在 src/engine，引擎脚本用 parents[2] 回推。
for (const relative of ['lib/engines/pptx-office/src/engine/pptx_bridge.py', 'lib/engines/pptx-office/assets/indigo-paperlight-v1.png', 'lib/engines/docx-complex-parse/docx_complex_parse.py']) {
  const path = join(INSTALLED, relative);
  console.log(`  ${existsSync(path) ? '✓' : '✗'} ${relative}${existsSync(path) ? `  ${statSync(path).size} B` : ''}`);
}

const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-0100', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 30000 });
const line = (handshake.stdout ?? '').split('\n').find((entry) => entry.trim()) ?? '';
try { console.log(`握手: ${JSON.stringify(JSON.parse(line).result?.serverInfo ?? {})}`); } catch { console.log(`握手: 无法解析 ${line.slice(0, 120)}`); }

const { resolveDshConfig } = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const config = await resolveDshConfig({ workspaceRoot: WORKSPACE });
console.log(`运行时解析: source=${config.env.DSH_DOCX_RUNTIME_SOURCE} 缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
console.log(`注入路径: DOCX_PYTHON=${config.env.DOCX_PYTHON ? '✓' : '✗'} DOCX_SOFFICE=${config.env.DOCX_SOFFICE ? '✓' : '✗'} DOCX_PDFTOPPM=${config.env.DOCX_PDFTOPPM ? '✓' : '✗'}`);

mkdirSync(WORKSPACE, { recursive: true });
const docxTarget = join(WORKSPACE, 'upstream-0100-source.docx');
copyFileSync(DOCX_SOURCE, docxTarget);

const child = spawn(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { env: { ...process.env, ...config.env }, stdio: ['pipe', 'pipe', 'pipe'] });
let buffer = ''; let id = 1; const pending = new Map();
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk;
  let index;
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
  if (reply.error) return { error: reply.error };
  const text = reply.result?.content?.find((part) => part.type === 'text')?.text ?? '';
  try { return JSON.parse(text); } catch { return { raw: text.slice(0, 400) }; }
};
const importPath = async (path) => {
  const result = await callTool('docx_import', { path });
  const match = JSON.stringify(result).match(/"id":"(sha256:[0-9a-f]+)"/);
  return { raw: result, id: match?.[1] };
};

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-0100', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);

const doctor = await callTool('docx_doctor', {});
console.log(`\ncapabilities: ${JSON.stringify(doctor.capabilities)}`);
console.log(`libreoffice : ${doctor.libreoffice?.available}  poppler: ${doctor.poppler?.available}  python: ${doctor.python?.available}`);

// 1) DOCX 链：写 + 渲染（渲染会走新 resolver 解析出来的 soffice/pdftoppm 路径）
const docx = await importPath(docxTarget);
console.log(`\n[1] docx_import -> ${docx.id}`);
if (docx.id) {
  const created = await callTool('docx_create', { requestId: 'up-0100-create', blueprint: { title: '上游 0.10.0 集成测试', body: ['装配 docx-runtime 后的端到端检查。'] } });
  const createdId = JSON.stringify(created).match(/"id":"(sha256:[0-9a-f]+)"/)?.[1];
  console.log(`    docx_create   : ${created.error ? `✗ ${JSON.stringify(created.error)}` : `✓ ${createdId}`}`);
  const rendered = await callTool('docx_render', { artifactRef: { id: docx.id, uri: 'managed' }, requestId: 'up-0100-render' });
  console.log(`    docx_render   : ${rendered.error ? `✗ ${JSON.stringify(rendered.error)}` : `✓ pages=${JSON.stringify(rendered.result?.pageCount ?? rendered.result?.pdf?.pageCount ?? '?')}`}`);
  const analyzed = await callTool('docx_analyze', { artifactRef: { id: docx.id, uri: 'managed' }, requestId: 'up-0100-analyze', complex: 'structure' });
  console.log(`    docx_analyze  : ${analyzed.error ? `✗ ${JSON.stringify(analyzed.error)}` : `✓ bridges=${analyzed.bridges?.length}`}`);
}

// 2) PPTX 新能力：applyArtStyle（要读打包进去的 indigo 封面位图）
const pptx = await importPath(PPTX_SOURCE);
console.log(`\n[2] pptx_import -> ${pptx.id}`);
if (pptx.id) {
  const styled = await callTool('pptx_call', { moduleId: 'pptx-office', operation: 'execute', requestId: 'up-0100-art', artifactRef: { id: pptx.id, uri: 'managed' }, payload: { action: 'applyArtStyle', styleId: 'indigo-paperlight-v1', artWord: 'DSH' } });
  console.log(`    applyArtStyle : ${styled.error ? `✗ ${JSON.stringify(styled.error)}` : `✓ styleId=${styled.result?.styleId} styledSlides=${styled.result?.styledSlides} artWord=${styled.result?.artWord}`}`);
  const verified = await callTool('pptx_call', { moduleId: 'pptx-office', operation: 'verify', requestId: 'up-0100-art-verify', artifactRef: { id: pptx.id, uri: 'managed' }, payload: { expectedSlideCount: 12 } });
  console.log(`    pptx_verify   : ${verified.error ? `✗ ${JSON.stringify(verified.error)}` : `✓ ok=${verified.ok ?? verified.result?.ok}`}`);
}

child.kill();
process.exit(0);

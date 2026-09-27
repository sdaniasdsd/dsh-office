// 把 0.9.3 的核心包装进 profile（运行时内容没变，保持已装好的那份即可），
// 然后回读：版本 / 握手 / 运行时解析 / doctor 能力表 / **在装好的布局里真跑一次
// docx-complex-parse**。最后一项是关键：live 会话里的服务进程要等 DSH 重启才会换核，
// 所以这里自己起一个服务，用装好的目录做端到端回执，不靠"应该会好"。
import { spawn, spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const CORE_TGZ = join(DIST, 'deepseek-ai-dsh-docx-0.9.3.tgz');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const SOURCE_DOCX = 'C:\\Users\\AA\\Desktop\\.dsh-docx\\objects\\dc8e138db5421b71df407d294199af3c8269597dbee6601aa9f84dafc69981c1\\artifact.docx';

if (!existsSync(CORE_TGZ)) throw new Error(`缺 tarball：${CORE_TGZ}`);
copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-0.9.3`);
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
const enginePy = join(INSTALLED, 'lib', 'engines', 'docx-complex-parse', 'docx_complex_parse.py');
console.log(`装好的引擎带 sys.path 修正: ${readFileSync(enginePy, 'utf8').includes('sys.path.insert')}`);

const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-093', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 30000 });
const line = (handshake.stdout ?? '').split('\n').find((entry) => entry.trim()) ?? '';
try { console.log(`握手: ${JSON.stringify(JSON.parse(line).result?.serverInfo ?? {})}`); } catch { console.log(`握手: 无法解析 ${line.slice(0, 100)}`); }

const { resolveDshConfig } = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const config = await resolveDshConfig({ workspaceRoot: WORKSPACE });
console.log(`运行时解析: source=${config.env.DSH_DOCX_RUNTIME_SOURCE} 缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);

// 工作区里放一份待解析文档，让服务通过它自己的 artifact store 收进去。
mkdirSync(WORKSPACE, { recursive: true });
const target = join(WORKSPACE, 'sweep-093-source.docx');
copyFileSync(SOURCE_DOCX, target);

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
  try { return JSON.parse(text); } catch { return { raw: text.slice(0, 300) }; }
};

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-093', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);

const report = await callTool('docx_doctor', {});
console.log(`capabilities: ${JSON.stringify(report.capabilities)}`);

const imported = await callTool('docx_import', { path: target });
const ref = imported.artifactRef ?? imported;
console.log(`导入: ${ref.id ?? JSON.stringify(imported).slice(0, 200)}`);

if (ref?.id) {
  const analyzed = await callTool('docx_analyze', { artifactRef: ref, requestId: 'install-093-analyze', complex: 'structure', easy: false });
  console.log(`docx_analyze(complex=structure): ${analyzed.error ? `✗ ${JSON.stringify(analyzed.error)}` : `✓ bridges=${analyzed.bridges?.length ?? '?'} complexMode=${analyzed.complexMode}`}`);
  const layout = await callTool('docx_analyze', { artifactRef: ref, requestId: 'install-093-layout', complex: 'layout', easy: false });
  console.log(`docx_analyze(complex=layout) : ${layout.error ? `✗ ${JSON.stringify(layout.error)}` : `✓ complexMode=${layout.complexMode}`}`);
}
child.kill();
process.exit(0);

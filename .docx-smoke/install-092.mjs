// 把 0.9.2 的核心包装进 profile（运行时内容没变，保持 0.9.1 那份已装好的即可），
// 然后回读版本 / 握手 / 运行时解析 / doctor 能力表。
import { spawn, spawnSync } from 'node:child_process';
import { copyFileSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const CORE_TGZ = join(DIST, 'deepseek-ai-dsh-docx-0.9.2.tgz');
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');

copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-0.9.2`);
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
const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-092', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 30000 });
const line = (handshake.stdout ?? '').split('\n').find((entry) => entry.trim()) ?? '';
try { console.log(`握手: ${JSON.stringify(JSON.parse(line).result?.serverInfo ?? {})}`); } catch { console.log(`握手: 无法解析 ${line.slice(0, 100)}`); }

// 用它自己的配置起服务，报 doctor
const { resolveDshConfig } = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const config = await resolveDshConfig({ workspaceRoot: 'C:\\Users\\AA\\Desktop\\bench' });
console.log(`运行时解析: source=${config.env.DSH_DOCX_RUNTIME_SOURCE} 缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);
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
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-092', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
const reply = await request('tools/call', { name: 'docx_doctor', arguments: {} });
const report = JSON.parse(reply.result.content.find((part) => part.type === 'text').text);
console.log(`capabilities: ${JSON.stringify(report.capabilities)}`);
console.log(`exceljs     : ${JSON.stringify(report.exceljs)}`);
child.kill();
process.exit(0);

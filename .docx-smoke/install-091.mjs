// 把 0.9.1 装进 DSH web profile：核心包 + 运行时包（与拆包后的真实用法一致），
// 然后逐项回读：版本、握手、doctor 能力表，并用**装好的那份**真跑一组调用。
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const CORE_TGZ = join(DIST, 'deepseek-ai-dsh-docx-0.9.1.tgz');
const RUNTIME_TGZ = join(DIST, 'deepseek-ai-dsh-docx-runtime-0.9.1.tgz');
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const INSTALLED_RUNTIME = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx-runtime');

for (const path of [CORE_TGZ, RUNTIME_TGZ]) if (!existsSync(path)) throw new Error(`缺资产：${path}`);

copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-0.9.1`);
const profile = JSON.parse(readFileSync(PROFILE_PACKAGE, 'utf8'));
const previous = profile.dependencies['@deepseek-ai/dsh-docx'];
profile.dependencies['@deepseek-ai/dsh-docx'] = `file:${CORE_TGZ.replaceAll('\\', '/')}`;
profile.dependencies['@deepseek-ai/dsh-docx-runtime'] = `file:${RUNTIME_TGZ.replaceAll('\\', '/')}`;
writeFileSync(PROFILE_PACKAGE, `${JSON.stringify(profile, null, 4)}\n`);
console.log(`profile 依赖：\n  @deepseek-ai/dsh-docx         was ${previous}\n                              now ${profile.dependencies['@deepseek-ai/dsh-docx']}\n  @deepseek-ai/dsh-docx-runtime now ${profile.dependencies['@deepseek-ai/dsh-docx-runtime']}`);
console.log(`备份：${PROFILE_PACKAGE}.before-0.9.1`);

for (let attempt = 1; attempt <= 2; attempt += 1) {
  console.log(`\n$ pnpm install --prefer-offline   (第 ${attempt} 次)`);
  const result = spawnSync('pnpm', ['install', '--prefer-offline'], { cwd: PROFILE, stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
  if (result.status === 0) break;
  console.log(`pnpm 退出码 ${result.status}`);
  if (attempt === 2) throw new Error('pnpm install 两次都没成');
}

const coreVersion = JSON.parse(readFileSync(join(INSTALLED, 'package.json'), 'utf8')).version;
const runtimeManifest = JSON.parse(readFileSync(join(INSTALLED_RUNTIME, 'runtime.json'), 'utf8'));
const server = readFileSync(join(INSTALLED, 'lib', 'server.mjs'), 'utf8');
console.log(`\n装的版本       : 核心 ${coreVersion}  运行时 ${runtimeManifest.version}（平台 ${runtimeManifest.platform}）`);
console.log(`运行时产地     : ${JSON.stringify(runtimeManifest.toolchain ?? null)}`);
console.log(`注册工具       : ${[...server.matchAll(/registerTool\(\s*["'`]([^"'`]+)["'`]/gu)].map((m) => m[1]).join(', ')}`);

// 用装好的这份起服务，做一组真实调用
const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'install-091', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 30000 });
const firstLine = (handshake.stdout ?? '').split('\n').find((line) => line.trim()) ?? '';
try { console.log(`握手           : ${JSON.stringify(JSON.parse(firstLine).result?.serverInfo ?? {})}`); } catch { console.log(`握手           : 无法解析 (${firstLine.slice(0, 120)})`); }

const runtimeConfig = await import(new URL(`file:///${join(INSTALLED, 'dsh', 'runtime-config.mjs').replaceAll('\\', '/')}`).href);
const config = await runtimeConfig.resolveDshConfig({ workspaceRoot: 'C:\\Users\\AA\\Desktop\\bench' });
console.log(`运行时解析     : source=${config.env.DSH_DOCX_RUNTIME_SOURCE}  缺失=${config.env.DSH_DOCX_RUNTIME_MISSING ?? '(无)'}`);

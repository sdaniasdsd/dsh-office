// Verify the packed 0.8.1 tarball as a consumer receives it, then install it into
// the DSH web profile and read the handshake back.
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const TGZ = 'D:\\开源团队作品\\dsh-office-releases\\v0.8.1\\deepseek-ai-dsh-docx-0.8.1.tgz';
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');
const INSTALLED = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');

const capture = (command, args) => {
  const result = spawnSync(command, args, { encoding: 'utf8', windowsHide: true, maxBuffer: 512 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`${command} ${args.slice(0, 3).join(' ')} failed: ${result.stderr}`);
  return result.stdout;
};

// 1. What the tarball itself contains.
const manifest = JSON.parse(capture('tar', ['-xOf', TGZ, 'package/package.json']));
const server = capture('tar', ['-xOf', TGZ, 'package/lib/server.mjs']);
console.log(`tarball      : ${TGZ}`);
console.log(`bytes        : ${statSync(TGZ).size}`);
console.log(`sha256       : ${createHash('sha256').update(readFileSync(TGZ)).digest('hex')}`);
console.log(`package      : ${manifest.name} ${manifest.version}`);
const construction = /new McpServer\(\{[^}]*\}/u.exec(server)?.[0] ?? '(not found)';
console.log(`McpServer    : ${construction}`);
console.log(`literal 0.5.0: ${server.split('0.5.0').length - 1}    literal 0.8.1: ${server.split('0.8.1').length - 1}`);
if (manifest.version !== '0.8.1') throw new Error('the tarball is not 0.8.1');
if (construction.includes('0.5.0')) throw new Error('the packaged bundle still hardcodes 0.5.0');

// 2. Install it into the web profile.
copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-0.8.1`);
const profile = JSON.parse(readFileSync(PROFILE_PACKAGE, 'utf8'));
const previous = profile.dependencies['@deepseek-ai/dsh-docx'];
profile.dependencies['@deepseek-ai/dsh-docx'] = `file:${TGZ.replaceAll('\\', '/')}`;
writeFileSync(PROFILE_PACKAGE, `${JSON.stringify(profile, null, 4)}\n`);
console.log(`\nprofile dependency:\n  was ${previous}\n  now ${profile.dependencies['@deepseek-ai/dsh-docx']}`);

for (let attempt = 1; attempt <= 2; attempt += 1) {
  console.log(`\n$ pnpm install --prefer-offline   (attempt ${attempt})`);
  const result = spawnSync('pnpm', ['install', '--prefer-offline'], { cwd: PROFILE, stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
  if (result.status === 0) break;
  console.log(`pnpm exited ${result.status}`);
  if (attempt === 2) throw new Error('pnpm install failed twice');
}

// 3. Read back what landed, including the handshake.
const installedManifest = JSON.parse(readFileSync(join(INSTALLED, 'package.json'), 'utf8'));
const installedServer = readFileSync(join(INSTALLED, 'lib', 'server.mjs'), 'utf8');
const tools = [...installedServer.matchAll(/registerTool\(\s*["'`]([^"'`]+)["'`]/gu)].map((match) => match[1]);
const initialize = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'verify-081', version: '1' } } });
const handshake = spawnSync(process.execPath, [join(INSTALLED, 'lib', 'server.mjs')], { input: `${initialize}\n`, encoding: 'utf8', windowsHide: true, timeout: 30000 });
const line = (handshake.stdout ?? '').split('\n').find((entry) => entry.trim().length > 0) ?? '';
let reported = '(no reply)';
try { reported = JSON.stringify(JSON.parse(line).result?.serverInfo ?? {}); } catch { reported = `(unparseable: ${line.slice(0, 100)})`; }
console.log(`\ninstalled version : ${installedManifest.version}`);
console.log(`tools             : ${tools.length} → ${tools.join(', ')}`);
console.log(`runtime           : ${['python', 'libreoffice', 'poppler'].filter((name) => existsSync(join(INSTALLED, 'runtime', 'win32-x64', name))).join(', ')}`);
console.log(`handshake         : ${reported}`);
console.log(`handshake matches : ${reported.includes(installedManifest.version)}`);

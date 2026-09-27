// Install a published DSH Office release into the web profile.
//
// Steps, each with its own receipt: read the release metadata from the API,
// download the asset, verify its digest against the one GitHub reports, look at
// what is inside the tarball, then point the profile at it and run pnpm install.
// Everything is spawned from Node so CJK paths never travel on a shell command
// line.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.8.0';
const ASSET = 'deepseek-ai-dsh-docx-0.8.0.tgz';
const DOWNLOAD_DIR = 'D:\\开源团队作品\\dsh-office-releases\\v0.8.0';
const TGZ = join(DOWNLOAD_DIR, ASSET);
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');

const run = (command, args, options = {}) => {
  const shown = args.map((arg) => (arg.length > 70 ? `${arg.slice(0, 67)}…` : arg)).join(' ');
  console.log(`\n$ ${command} ${shown}`);
  const result = spawnSync(command, args, { stdio: 'inherit', windowsHide: true, ...options });
  if (result.status !== 0) throw new Error(`${command} exited ${result.status}`);
  return result;
};
const capture = (command, args) => {
  const result = spawnSync(command, args, { encoding: 'utf8', windowsHide: true, maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed: ${result.stderr}`);
  return result.stdout;
};

// 1. What the release says about itself.
const release = JSON.parse(capture('gh', ['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,name,isDraft,isPrerelease,publishedAt,assets']));
console.log(`release    : ${release.tagName}  "${release.name}"  published ${release.publishedAt}  draft=${release.isDraft} prerelease=${release.isPrerelease}`);
const asset = release.assets.find((entry) => entry.name === ASSET);
if (!asset) throw new Error(`${ASSET} is not attached to ${TAG}`);
console.log(`asset      : ${asset.name}  ${asset.size} bytes  ${asset.label ?? ''}`);
console.log(`digest     : ${asset.digest}`);
console.log(`url        : ${asset.url}`);

// 2. Download, unless a verified copy is already here.
mkdirSync(DOWNLOAD_DIR, { recursive: true });
const sha256 = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');
const expected = String(asset.digest).replace(/^sha256:/u, '');
const localMatches = existsSync(TGZ) && statSync(TGZ).size === asset.size && sha256(TGZ) === expected;
if (localMatches) {
  console.log(`\nlocal copy already matches the published digest: ${TGZ}`);
} else {
  console.log(`\ndownloading to ${TGZ}`);
  run('gh', ['release', 'download', TAG, '--repo', REPO, '--pattern', ASSET, '--dir', DOWNLOAD_DIR, '--clobber']);
}
const size = statSync(TGZ).size;
const digest = sha256(TGZ);
console.log(`\nlocal file : ${TGZ}`);
console.log(`bytes      : ${size} (published ${asset.size})`);
console.log(`sha256     : ${digest}`);
console.log(`matches    : ${digest === expected && size === asset.size}`);
if (digest !== expected || size !== asset.size) throw new Error('the downloaded package does not match the published digest');

// 3. What is inside the tarball: npm packs under a `package/` prefix.
const manifest = JSON.parse(capture('tar', ['-xOf', TGZ, 'package/package.json']));
console.log(`\npackage    : ${manifest.name} ${manifest.version}`);
console.log(`bundles    : ${manifest.dsh?.bundle?.patch ?? '(none)'}`);

// 4. Point the profile at it and install.
copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-${manifest.version}`);
const profile = JSON.parse(readFileSync(PROFILE_PACKAGE, 'utf8'));
const previous = profile.dependencies['@deepseek-ai/dsh-docx'];
profile.dependencies['@deepseek-ai/dsh-docx'] = `file:${TGZ.replaceAll('\\', '/')}`;
writeFileSync(PROFILE_PACKAGE, `${JSON.stringify(profile, null, 4)}\n`);
console.log(`\nprofile dependency:\n  was ${previous}\n  now ${profile.dependencies['@deepseek-ai/dsh-docx']}`);
console.log(`backup: ${PROFILE_PACKAGE}.before-${manifest.version}`);
run('pnpm', ['install', '--prefer-offline'], { cwd: PROFILE });

// 5. Verify what actually landed.
const installed = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const installedManifest = JSON.parse(readFileSync(join(installed, 'package.json'), 'utf8'));
const server = readFileSync(join(installed, 'lib', 'server.mjs'), 'utf8');
const tools = [...server.matchAll(/registerTool\('([^']+)'/gu)].map((match) => match[1]);
const handshake = JSON.parse(capture(process.execPath, [join(installed, 'lib', 'server.mjs')]).split('\n')[0] ?? '{}');
console.log(`\ninstalled version : ${installedManifest.version}`);
console.log(`registered tools  : ${tools.join(', ')}`);
console.log(`format work present: cellFormats=${(server.match(/cellFormats/gu) ?? []).length} firstLineIndentPt=${(server.match(/firstLineIndentPt/gu) ?? []).length} parseFormatting=${(server.match(/parseFormatting/gu) ?? []).length}`);
console.log(`runtime present   : ${existsSync(join(installed, 'runtime', 'win32-x64', 'python', 'python.exe'))}`);
console.log(`runtime dirs      : ${['python', 'libreoffice', 'poppler'].filter((name) => existsSync(join(installed, 'runtime', 'win32-x64', name))).join(', ')}`);
console.log(`handshake         : ${handshake.result?.serverInfo ? JSON.stringify(handshake.result.serverInfo) : '(no reply on an empty stdin)'}`);

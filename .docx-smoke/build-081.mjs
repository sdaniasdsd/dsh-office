// Build 0.8.1 from the fix commit in a clean worktree, give it the runtime, and
// pack it.
//
// Why a worktree: the clone's working tree carries another line's uncommitted
// work, so a build there would ship code the repository does not have. The
// runtime is copied from the installed 0.8.0 package, which is exactly what the
// published 0.8.0 carries and is not being rewritten by anyone.
import { spawnSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const WORKTREE = 'D:\\认真版agent\\.docx-smoke\\build-worktree\\the-last-docx';
const DIST = join(WORKTREE, 'dist', 'dsh-docx');
const INSTALLED = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
const OUT_DIR = 'D:\\开源团队作品\\dsh-office-releases\\v0.8.1';

const run = (command, args, options = {}) => {
  console.log(`\n$ ${command} ${args.join(' ')}   (cwd ${options.cwd ?? process.cwd()})`);
  const result = spawnSync(command, args, { stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32', ...options });
  if (result.status !== 0) throw new Error(`${command} exited ${result.status}`);
};

// 1. Bundle from the committed sources. `runtime/` is not in the repository (it
// is where the downloaded runtimes live), so a fresh checkout has no such
// directory and build-dsh.mjs fails on the last copy with ENOENT — create it.
mkdirSync(join(WORKTREE, 'runtime'), { recursive: true });
run('node', ['scripts/build-dsh.mjs'], { cwd: WORKTREE });
const manifest = JSON.parse(readFileSync(join(DIST, 'package.json'), 'utf8'));
console.log(`\nbuilt ${manifest.name} ${manifest.version}`);
if (manifest.version !== '0.8.1') throw new Error(`expected 0.8.1, built ${manifest.version}`);

// 2. The handshake version must now come from the injected constant.
const server = readFileSync(join(DIST, 'lib', 'server.mjs'), 'utf8');
const construction = /new McpServer\(\{[^}]*\}/u.exec(server)?.[0] ?? '(not found)';
console.log(`McpServer    : ${construction}`);
console.log(`literal 0.5.0: ${server.split('0.5.0').length - 1}`);
console.log(`literal 0.8.1: ${server.split('0.8.1').length - 1}`);
if (construction.includes('0.5.0')) throw new Error('the built bundle still hardcodes 0.5.0');

// 3. Runtime: same bytes the published 0.8.0 ships.
const runtimeTarget = join(DIST, 'runtime');
console.log(`\ncopying the runtime from the installed 0.8.0 package into the build…`);
rmSync(runtimeTarget, { recursive: true, force: true });
cpSync(join(INSTALLED, 'runtime'), runtimeTarget, { recursive: true });
const runtimeFiles = statSync(join(runtimeTarget, 'win32-x64')).isDirectory();
console.log(`runtime present: ${runtimeFiles}`);
for (const name of ['python', 'libreoffice', 'poppler']) {
  console.log(`  ${name.padEnd(11)}: ${existsSync(join(runtimeTarget, 'win32-x64', name)) ? 'ok' : 'MISSING'}`);
}

// 4. Pack, and move the tarball next to the other releases.
run('npm', ['pack'], { cwd: DIST });
const produced = join(DIST, `deepseek-ai-dsh-docx-${manifest.version}.tgz`);
if (!existsSync(produced)) throw new Error(`npm pack did not produce ${produced}`);
mkdirSync(OUT_DIR, { recursive: true });
const tgz = join(OUT_DIR, `deepseek-ai-dsh-docx-${manifest.version}.tgz`);
rmSync(tgz, { force: true });
renameSync(produced, tgz);
const bytes = statSync(tgz).size;
const sha256 = createHash('sha256').update(readFileSync(tgz)).digest('hex');
console.log(`\ntarball: ${tgz}`);
console.log(`bytes  : ${bytes}  (${(bytes / 1048576).toFixed(1)} MB)`);
console.log(`sha256 : ${sha256}`);
console.log(`\nnext: gh release create v${manifest.version} <tarball> --repo sdaniasdsd/dsh-office --target main`);

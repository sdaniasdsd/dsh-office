// Pack the freshly built plugin and install it into the DSH web profile.
//
// Two rules learned the hard way: every path that contains CJK characters is
// either a working directory or comes out of a file, never a command-line
// argument (the Windows shell mangles those); and the working copy of the
// profile's package.json is backed up before it is edited.
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'D:\\开源团队作品\\the-last-docx\\dist';
const PACKAGE_DIR = join(DIST, 'dsh-docx');
const VERSION = JSON.parse(readFileSync(join(PACKAGE_DIR, 'package.json'), 'utf8')).version;
const TGZ_NAME = `deepseek-ai-dsh-docx-${VERSION}.tgz`;
const TGZ = join(DIST, TGZ_NAME);
const PROFILE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web');
const PROFILE_PACKAGE = join(PROFILE, 'package.json');

const run = (command, args, cwd) => {
  console.log(`\n$ ${command} ${args.join(' ')}   (cwd ${cwd})`);
  const result = spawnSync(command, args, { cwd, stdio: 'inherit', shell: process.platform === 'win32' });
  if (result.status !== 0) {
    console.error(`FAILED (exit ${result.status})`);
    process.exit(result.status ?? 1);
  }
};

// Pack with the package directory as the working directory, so no CJK path is
// ever an argument; the tarball is then moved into dist by the filesystem API.
run('npm', ['pack'], PACKAGE_DIR);
const produced = join(PACKAGE_DIR, TGZ_NAME);
if (!existsSync(produced)) {
  console.error(`expected ${produced} — npm pack produced a different name`);
  process.exit(1);
}
rmSync(TGZ, { force: true });
copyFileSync(produced, TGZ);
rmSync(produced, { force: true });
console.log(`\ntarball: ${TGZ}  ${(statSync(TGZ).size / 1048576).toFixed(0)} MB`);

copyFileSync(PROFILE_PACKAGE, `${PROFILE_PACKAGE}.before-${VERSION}`);
const profile = JSON.parse(readFileSync(PROFILE_PACKAGE, 'utf8'));
const previous = profile.dependencies['@deepseek-ai/dsh-docx'];
profile.dependencies['@deepseek-ai/dsh-docx'] = `file:${TGZ.replaceAll('\\', '/')}`;
writeFileSync(PROFILE_PACKAGE, `${JSON.stringify(profile, null, 4)}\n`);
console.log(`\nprofile dependency:\n  was ${previous}\n  now ${profile.dependencies['@deepseek-ai/dsh-docx']}`);
console.log(`backup: ${PROFILE_PACKAGE}.before-${VERSION}`);

run('pnpm', ['install', '--prefer-offline'], PROFILE);

const installed = join(PROFILE, 'node_modules', '@deepseek-ai', 'dsh-docx');
const installedVersion = JSON.parse(readFileSync(join(installed, 'package.json'), 'utf8')).version;
const server = readFileSync(join(installed, 'lib', 'server.mjs'), 'utf8');
console.log(`\ninstalled version        : ${installedVersion}`);
console.log(`cellFormats in bundle    : ${(server.match(/cellFormats/gu) ?? []).length}`);
console.log(`parseFormatting in bundle: ${(server.match(/parseFormatting/gu) ?? []).length}`);
console.log(`runtime present          : ${existsSync(join(installed, 'runtime', 'win32-x64', 'python', 'python.exe'))}`);
console.log(`runtime size             : ${(statSync(join(installed, 'runtime', 'win32-x64', 'libreoffice', 'program')).isDirectory() ? 'libreoffice ok' : 'missing')}`);

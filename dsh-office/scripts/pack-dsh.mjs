// 打包三个包：core / runtime / full。
// 运行时不存在的构建里只会产出 core（build-dsh.mjs 已经打印过原因），这里按存在的目录打包。
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const dist = join(root, 'dist');

const names = ['dsh-docx', 'dsh-docx-runtime', 'dsh-docx-full'].filter((name) => existsSync(join(dist, name, 'package.json')));
if (!names.length) {
  console.error('no built packages under dist/ — run `npm run build:dsh` first');
  process.exit(1);
}
for (const name of names) {
  const manifest = JSON.parse(readFileSync(join(dist, name, 'package.json'), 'utf8'));
  const tarball = `${manifest.name.replace('@', '').replace('/', '-')}-${manifest.version}.tgz`;
  console.log(`\n$ npm pack ./dist/${name}   → ${tarball}   (${manifest.name} ${manifest.version})`);
  const result = spawnSync('npm', ['pack', join(dist, name), '--pack-destination', dist], { stdio: 'inherit', windowsHide: true, shell: process.platform === 'win32' });
  if (result.status !== 0) process.exit(result.status ?? 1);
}
console.log(`\npacked ${names.length} package(s) into ${dist}`);

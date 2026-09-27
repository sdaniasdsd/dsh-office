// 从 dsh-toolchain 取运行时：按 toolchain.lock.json 钉住的 tag/资产/摘要下载并校验，
// 解到 runtime/win32-x64/（build-dsh.mjs 认这个位置），最后逐项确认三件二进制能用。
//
//   node scripts/fetch-toolchain.mjs                 下载 + 校验 + 解包 + 自检
//   node scripts/fetch-toolchain.mjs --check-only     只做自检（运行时已就位时用）
//
// 为什么用 gh / Node 而不是 PowerShell 脚本：路径与资产名里可能有非 ASCII，命令行传参会出问题；
// 下载与校验都在 Node 里做，参数不经过 shell。
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const pin = JSON.parse(readFileSync(join(root, 'toolchain.lock.json'), 'utf8'));
const runtimeRoot = join(root, 'runtime', pin.platform);
const cacheDir = join(root, '.build-cache', 'toolchain');
const checkOnly = process.argv.includes('--check-only');

const binaries = [
  { name: 'python', path: join(runtimeRoot, 'python', 'python.exe'), args: ['--version'] },
  { name: 'libreoffice', path: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'), args: ['--version'], timeout: 60000 },
  { name: 'poppler', path: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'), args: ['-v'] },
];

function selfCheck() {
  let ok = true;
  for (const binary of binaries) {
    if (!existsSync(binary.path)) { console.log(`  ✗ ${binary.name.padEnd(12)} 缺失 ${binary.path}`); ok = false; continue; }
    const probe = spawnSync(binary.path, binary.args, { encoding: 'utf8', windowsHide: true, timeout: binary.timeout ?? 30000 });
    const output = `${probe.stdout ?? ''}${probe.stderr ?? ''}`.trim().split('\n')[0] ?? '';
    const good = probe.status === 0 && Boolean(output);
    if (!good) ok = false;
    console.log(`  ${good ? '✓' : '✗'} ${binary.name.padEnd(12)} ${output || `exit ${probe.status}`}`);
  }
  return ok;
}

console.log(`工具链钉住：${pin.repo} ${pin.tag} → ${pin.asset}`);
console.log(`期望 sha256：${pin.sha256}`);
console.log(`目标目录  ：${runtimeRoot}`);

if (checkOnly) {
  console.log('\n自检（--check-only）：');
  process.exit(selfCheck() ? 0 : 1);
}

mkdirSync(cacheDir, { recursive: true });
const tarball = join(cacheDir, pin.asset);
const sha256 = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');

if (existsSync(tarball) && statSync(tarball).size === pin.bytes && sha256(tarball) === pin.sha256) {
  console.log(`\n缓存命中：${tarball}（字节数与 sha256 与钉住的一致）`);
} else {
  rmSync(tarball, { force: true });
  console.log(`\n下载 ${pin.asset} …`);
  const download = spawnSync('gh', ['release', 'download', pin.tag, '--repo', pin.repo, '--pattern', pin.asset, '--dir', cacheDir, '--clobber'], { stdio: 'inherit', windowsHide: true });
  if (download.status !== 0) {
    console.error('gh release download 失败（需要已登录的 gh，或网络可达）。');
    process.exit(download.status ?? 1);
  }
}
const bytes = statSync(tarball).size;
const digest = sha256(tarball);
console.log(`\n下载完成：${bytes} B（钉住 ${pin.bytes}）`);
console.log(`sha256  ：${digest}`);
if (bytes !== pin.bytes || digest !== pin.sha256) {
  console.error('校验失败：字节数或 sha256 与 toolchain.lock.json 不一致，拒绝解包。');
  process.exit(1);
}
console.log('校验通过 ✓');

// npm 打包的 tgz 里顶层是 package/，运行时在 package/runtime/win32-x64/ 下。
// strip 掉 package/runtime 两级后，剩下的 win32-x64/... 正好落到 <root>/runtime/ 里。
console.log(`\n解包到 ${runtimeRoot} …`);
rmSync(runtimeRoot, { recursive: true, force: true });
mkdirSync(join(root, 'runtime'), { recursive: true });
const extract = spawnSync('tar', ['-xzf', tarball, '-C', join(root, 'runtime'), '--strip-components=2', `package/runtime/${pin.platform}`], { stdio: 'inherit', windowsHide: true });
if (extract.status !== 0 || !existsSync(join(runtimeRoot, 'python', 'python.exe'))) {
  console.error(`tar 解包失败（exit ${extract.status}）或解出来的目录结构不符（缺 ${join(runtimeRoot, 'python', 'python.exe')}）`);
  process.exit(extract.status ?? 1);
}
console.log(`解包完成：${runtimeRoot}`);

console.log('\n自检：');
const ok = selfCheck();
console.log(ok ? '\n运行时就绪：可以跑 npm run build:dsh / pack:dsh 了' : '\n运行时自检未全部通过，先别打包');
process.exit(ok ? 0 : 1);

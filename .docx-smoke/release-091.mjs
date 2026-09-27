// 发布 0.9.1：三个资产一次传（元包依赖这两个子包的直链，先传子包再传元包是同一次 gh 调用内的顺序）。
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.9.1';
const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const NAMES = ['deepseek-ai-dsh-docx-0.9.1.tgz', 'deepseek-ai-dsh-docx-runtime-0.9.1.tgz', 'deepseek-ai-dsh-docx-full-0.9.1.tgz'];
const ASSETS = NAMES.map((name) => join(DIST, name));
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes-091.md';

const gh = (args, capture = false) => {
  const result = spawnSync('gh', args, { encoding: 'utf8', windowsHide: true, stdio: capture ? 'pipe' : 'inherit', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`gh ${args[0] ?? ''} exited ${result.status}: ${result.stderr ?? ''}`);
  return result.stdout;
};

const local = new Map();
for (const path of ASSETS) {
  const digest = createHash('sha256').update(readFileSync(path)).digest('hex');
  local.set(path.split('\\').pop(), `sha256:${digest}`);
  console.log(`本地 ${path.split('\\').pop().padEnd(42)} ${String(statSync(path).size).padStart(12)} B  ${digest}`);
}

gh(['release', 'create', TAG, ...ASSETS, '--repo', REPO, '--target', 'main',
  '--title', 'DSH Office 0.9.1 — 适配工具链交接（运行时改为从 dsh-toolchain 取）',
  '--notes-file', NOTES]);

const release = JSON.parse(gh(['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,isDraft,isPrerelease,url,assets'], true));
console.log(`\ntag  : ${release.tagName}  draft=${release.isDraft} prerelease=${release.isPrerelease}`);
console.log(`url  : ${release.url}`);
let ok = true;
for (const asset of release.assets) {
  const matches = asset.digest === local.get(asset.name);
  if (!matches) ok = false;
  console.log(`asset: ${asset.name}  ${asset.size} B  state=${asset.state}  ${matches ? 'digest 一致 ✓' : `digest 不一致 ✗ ${asset.digest}`}`);
}
console.log(`\n全部一致: ${ok}`);
process.exit(ok ? 0 : 1);

// 发布 v0.9.0：把拆出来的三个包作为同一个 release 的三个资产传上去。
// 全程用 gh spawn（CJK 路径不经 shell），上传耗时放后台。
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.9.0';
const DIST = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist';
const ASSETS = [
  'deepseek-ai-dsh-docx-0.9.0.tgz',
  'deepseek-ai-dsh-docx-runtime-0.9.0.tgz',
  'deepseek-ai-dsh-docx-full-0.9.0.tgz',
].map((name) => join(DIST, name));
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes-090.md';

const gh = (args, capture = false) => {
  const result = spawnSync('gh', args, { encoding: 'utf8', windowsHide: true, stdio: capture ? 'pipe' : 'inherit', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`gh ${args[0] ?? ''} exited ${result.status}: ${result.stderr ?? ''}`);
  return result.stdout;
};

console.log('本地资产：');
const local = new Map();
for (const path of ASSETS) {
  const bytes = statSync(path).size;
  const digest = createHash('sha256').update(readFileSync(path)).digest('hex');
  local.set(path.split('\\').pop(), `sha256:${digest}`);
  console.log(`  ${path.split('\\').pop().padEnd(42)} ${String(bytes).padStart(12)} B  ${digest}`);
}

gh(['release', 'create', TAG, ...ASSETS, '--repo', REPO, '--target', 'main',
  '--title', 'DSH Office 0.9.0 — 运行时独立成包（核心 2.2 MB / 运行时 546 MB / 元包 651 B）',
  '--notes-file', NOTES]);

const release = JSON.parse(gh(['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,isDraft,isPrerelease,url,assets'], true));
console.log(`\ntag  : ${release.tagName}  draft=${release.isDraft} prerelease=${release.isPrerelease}`);
console.log(`url  : ${release.url}`);
let allMatch = true;
for (const asset of release.assets) {
  const expected = local.get(asset.name);
  const matches = asset.digest === expected;
  if (!matches) allMatch = false;
  console.log(`asset: ${asset.name}`);
  console.log(`  ${asset.size} B  state=${asset.state}`);
  console.log(`  digest ${asset.digest}  ${matches ? '= 本地一致 ✓' : `≠ 本地 ${expected} ✗`}`);
}
console.log(`\n全部 digest 与本地一致: ${allMatch}`);
process.exit(allMatch ? 0 : 1);

// 在 dsh-toolchain 里建 v0.9.0 release 并挂上运行时交接包（546 MB，耗时放后台）。
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';

const REPO = 'sdaniasdsd/dsh-toolchain';
const TAG = 'v0.9.0';
const ASSET = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist\\deepseek-ai-dsh-docx-runtime-0.9.0.tgz';
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes-toolchain-090.md';

const gh = (args, capture = false) => {
  const result = spawnSync('gh', args, { encoding: 'utf8', windowsHide: true, stdio: capture ? 'pipe' : 'inherit', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`gh ${args[0] ?? ''} exited ${result.status}: ${result.stderr ?? ''}`);
  return result.stdout;
};

const bytes = statSync(ASSET).size;
const digest = createHash('sha256').update(readFileSync(ASSET)).digest('hex');
console.log(`本地资产: ${ASSET.split('\\').pop()}  ${bytes} B`);
console.log(`sha256  : ${digest}`);

gh(['release', 'create', TAG, ASSET, '--repo', REPO, '--target', 'main',
  '--title', 'DSH Office 运行时 0.9.0（Windows x64：Python 3.13 + LibreOffice 26.8.0 + Poppler 26.09.0）',
  '--notes-file', NOTES]);

const release = JSON.parse(gh(['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,isDraft,isPrerelease,url,assets'], true));
console.log(`\ntag  : ${release.tagName}  draft=${release.isDraft} prerelease=${release.isPrerelease}`);
console.log(`url  : ${release.url}`);
let ok = true;
for (const asset of release.assets) {
  const matches = asset.digest === `sha256:${digest}`;
  if (!matches) ok = false;
  console.log(`asset: ${asset.name}  ${asset.size} B  state=${asset.state}`);
  console.log(`  digest ${asset.digest}  ${matches ? '= 本地一致 ✓' : '≠ 本地 ✗'}`);
}
console.log(`\ndigest 与本地一致: ${ok}`);
process.exit(ok ? 0 : 1);

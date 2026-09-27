// 发 v0.9.3：只附核心包（引擎 sys.path 修正在这里；运行时内容未变，仍用 0.9.1 那份）。
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.9.3';
const ASSET = join('D:\\开源团队作品\\dsh-office-clone\\dsh-office\\dist', 'deepseek-ai-dsh-docx-0.9.3.tgz');
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes-093.md';

const gh = (args, capture = false) => {
  const result = spawnSync('gh', args, { encoding: 'utf8', windowsHide: true, stdio: capture ? 'pipe' : 'inherit', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`gh ${args[0] ?? ''} exited ${result.status}: ${result.stderr ?? ''}`);
  return result.stdout;
};

const digest = createHash('sha256').update(readFileSync(ASSET)).digest('hex');
console.log(`资产 ${ASSET.split('\\').pop()}  ${statSync(ASSET).size} B  ${digest}`);

gh(['release', 'create', TAG, ASSET, '--repo', REPO, '--target', 'main',
  '--title', 'DSH Office 0.9.3 — 修好随包运行时下必然崩溃的 docx-complex-parse',
  '--notes-file', NOTES]);

const release = JSON.parse(gh(['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,isDraft,url,assets'], true));
console.log(`\ntag ${release.tagName}  draft=${release.isDraft}  ${release.url}`);
for (const asset of release.assets) {
  console.log(`  ${asset.name}  ${asset.size} B  ${asset.state}  ${asset.digest === `sha256:${digest}` ? 'digest 一致 ✓' : `digest ✗ ${asset.digest}`}`);
}

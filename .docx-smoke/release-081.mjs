// Create the v0.8.1 release with the locally built package as its asset.
import { spawnSync } from 'node:child_process';
import { statSync } from 'node:fs';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.8.1';
const ASSET = 'D:\\开源团队作品\\dsh-office-releases\\v0.8.1\\deepseek-ai-dsh-docx-0.8.1.tgz';
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes-081.md';

const gh = (args, capture = false) => {
  const result = spawnSync('gh', args, { encoding: 'utf8', windowsHide: true, stdio: capture ? 'pipe' : 'inherit', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`gh ${args[0] ?? ''} exited ${result.status}: ${result.stderr ?? ''}`);
  return result.stdout;
};

console.log(`asset: ${ASSET}  ${(statSync(ASSET).size / 1048576).toFixed(1)} MB`);
gh(['release', 'create', TAG, ASSET, '--repo', REPO, '--target', 'main',
  '--title', 'dsh-docx 0.8.1 — the MCP handshake reports the packaged version',
  '--notes-file', NOTES]);

const release = JSON.parse(gh(['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,isDraft,isPrerelease,url,assets'], true));
console.log(`\ntag  : ${release.tagName}  draft=${release.isDraft} prerelease=${release.isPrerelease}`);
console.log(`url  : ${release.url}`);
for (const asset of release.assets) {
  console.log(`asset: ${asset.name}  ${asset.size} bytes  state=${asset.state}  digest=${asset.digest ?? '(none)'}`);
}

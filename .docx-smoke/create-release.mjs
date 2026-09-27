// Create the GitHub release and attach the 539 MB package.
//
// gh is spawned from Node so the CJK paths travel through the wide-character API
// instead of a shell command line, and the notes go in through --notes-file
// (a UTF-8 file) rather than as an argument.
import { spawnSync } from 'node:child_process';
import { statSync } from 'node:fs';

const REPO = 'sdaniasdsd/dsh-office';
const TAG = 'v0.7.0';
const ASSET = 'D:\\开源团队作品\\the-last-docx\\dist\\deepseek-ai-dsh-docx-0.7.0.tgz';
const NOTES = 'D:\\认真版agent\\.docx-smoke\\release-notes.md';

const gh = (args) => {
  console.log(`\n$ gh ${args.map((arg) => (arg.length > 60 ? `${arg.slice(0, 57)}…` : arg)).join(' ')}`);
  const result = spawnSync('gh', args, { stdio: 'inherit', windowsHide: true });
  if (result.status !== 0) throw new Error(`gh exited ${result.status}`);
};

console.log(`asset: ${ASSET}  ${(statSync(ASSET).size / 1048576).toFixed(1)} MB`);
gh(['release', 'create', TAG, ASSET, '--repo', REPO, '--target', 'main',
  '--title', 'dsh-docx 0.7.0 — reference rebuilds now carry run, cell and indent formatting',
  '--notes-file', NOTES]);

const view = spawnSync('gh', ['release', 'view', TAG, '--repo', REPO, '--json', 'tagName,name,assets,url'], { encoding: 'utf8', windowsHide: true });
if (view.status !== 0) throw new Error(`release view failed: ${view.stderr}`);
const release = JSON.parse(view.stdout);
console.log(`\ntag    : ${release.tagName}`);
console.log(`url    : ${release.url}`);
for (const asset of release.assets) {
  console.log(`asset  : ${asset.name}  ${(asset.size / 1048576).toFixed(1)} MB  state=${asset.state}`);
}

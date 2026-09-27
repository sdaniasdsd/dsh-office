// 用 gh 读两个仓库的现状（Node 传参，避免 PowerShell 把 --jq 里的空格拆散）。
import { execFileSync } from 'node:child_process';

const api = (path) => JSON.parse(execFileSync('gh', ['api', path], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }));

console.log('=== dsh-toolchain 远端 tree ===');
const tree = api('repos/sdaniasdsd/dsh-toolchain/git/trees/main?recursive=1');
for (const entry of tree.tree.filter((item) => item.type === 'blob')) console.log(`  ${entry.path}  ${entry.size}B`);
const head = api('repos/sdaniasdsd/dsh-toolchain/commits/main');
console.log(`  HEAD ${head.sha.slice(0, 12)}  ${head.commit.message.split('\n')[0]}`);

for (const repo of ['sdaniasdsd/dsh-office', 'sdaniasdsd/dsh-toolchain']) {
  console.log(`\n=== ${repo} releases ===`);
  const releases = api(`repos/${repo}/releases`);
  if (!releases.length) { console.log('  (无 release)'); continue; }
  for (const release of releases) {
    console.log(`  ${release.tag_name}  ${release.published_at}  draft=${release.draft}`);
    for (const asset of release.assets) {
      console.log(`    ${asset.name}  ${asset.size} B  state=${asset.state}  ${asset.digest ?? '(no digest)'}`);
    }
  }
}

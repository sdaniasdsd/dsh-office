// 看 dsh-toolchain 现在到底长什么样：谁提交的、有没有新提交/分支/release、文件树是什么。
// 也顺手列出账号下最近更新的仓库，确认"工具链已经完成"指的是哪个仓库。
import { execFileSync } from 'node:child_process';

const api = (path) => JSON.parse(execFileSync('gh', ['api', path], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }));

console.log('=== 账号下最近更新的仓库 ===');
for (const repo of api('users/sdaniasdsd/repos?sort=updated&per_page=10')) {
  console.log(`  ${repo.name.padEnd(24)} ${repo.visibility.padEnd(7)} pushed=${repo.pushed_at}  size=${repo.size}KB  ${repo.description ?? ''}`);
}

console.log('\n=== dsh-toolchain 提交历史 ===');
for (const commit of api('repos/sdaniasdsd/dsh-toolchain/commits?per_page=10')) {
  console.log(`  ${commit.sha.slice(0, 10)}  ${commit.commit.author.date}  ${commit.commit.author.name}  ${commit.commit.message.split('\n')[0]}`);
}

console.log('\n=== dsh-toolchain 分支 / release ===');
for (const branch of api('repos/sdaniasdsd/dsh-toolchain/branches')) console.log(`  branch ${branch.name} ${branch.commit.sha.slice(0, 10)}`);
for (const release of api('repos/sdaniasdsd/dsh-toolchain/releases')) {
  console.log(`  release ${release.tag_name} published=${release.published_at} draft=${release.draft}`);
  for (const asset of release.assets) console.log(`    ${asset.name} ${asset.size}B ${asset.state}`);
}

console.log('\n=== dsh-toolchain 文件树（main） ===');
const tree = api('repos/sdaniasdsd/dsh-toolchain/git/trees/main?recursive=1');
for (const entry of tree.tree.filter((item) => item.type === 'blob')) console.log(`  ${entry.path}  ${entry.size}B`);

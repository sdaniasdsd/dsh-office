// 只读 v0.9.1 的资产与说明首段，确认 -full 已撤下、说明已更正。
import { execFileSync } from 'node:child_process';

const release = JSON.parse(execFileSync('gh', ['api', 'repos/sdaniasdsd/dsh-office/releases/tags/v0.9.1'], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }));
console.log(`tag ${release.tag_name}  published=${release.published_at}  draft=${release.draft}`);
console.log(`assets: ${release.assets.length}`);
for (const asset of release.assets) console.log(`  ${asset.name}  ${asset.size} B  ${asset.state}  ${asset.digest}`);
const body = release.body ?? '';
const correction = body.split('\n').filter((line) => line.includes('EXOTIC') || line.includes('撤下') || line.includes('两步')).slice(0, 4);
console.log('\n说明里与更正相关的行：');
for (const line of correction) console.log(`  ${line.trim().slice(0, 140)}`);
console.log(`\n说明总长度 ${body.length} 字符`);

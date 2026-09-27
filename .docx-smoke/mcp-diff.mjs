// What exactly is different between my src/mcp.ts and the one committed in the
// publishing clone? Printed as a line diff, since the two files are one long line
// per tool registration.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const CLONE_ROOT = 'D:\\开源团队作品\\dsh-office-clone';
const MINE = 'D:\\开源团队作品\\the-last-docx\\src\\mcp.ts';

const head = spawnSync('git', ['-C', CLONE_ROOT, 'show', 'HEAD:the-last-docx/src/mcp.ts'], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
if (head.status !== 0) throw new Error(head.stderr);
const committed = head.stdout;
const mine = readFileSync(MINE, 'utf8');

if (committed === mine) {
  console.log('identical — nothing to port');
  process.exit(0);
}
const a = committed.split('\n');
const b = mine.split('\n');
console.log(`committed ${a.length} lines / ${committed.length} chars; mine ${b.length} lines / ${mine.length} chars`);
const max = Math.max(a.length, b.length);
for (let index = 0; index < max; index += 1) {
  if (a[index] === b[index]) continue;
  console.log(`\n--- line ${index + 1} ---`);
  console.log(`  committed: ${String(a[index] ?? '(absent)').slice(0, 300)}`);
  console.log(`  mine     : ${String(b[index] ?? '(absent)').slice(0, 300)}`);
  console.log(`  committed len ${String(a[index] ?? '').length} / mine len ${String(b[index] ?? '').length}`);
}

// Fix the escape that broke the committed blob, in the index only.
//
// The committed src/mcp.ts carries `reference's` inside a single-quoted tool
// description, so the file does not parse; the working file has `reference\'s`.
// Typechecking the commit (not the working tree) is what caught it.
import { spawnSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';

const CLONE_ROOT = 'D:\\开源团队作品\\dsh-office-clone';
const PATH_IN_REPO = 'the-last-docx/src/mcp.ts';
const TEMP = join('D:\\认真版agent\\.docx-smoke', 'mcp-staged.ts');

const git = (args) => {
  const result = spawnSync('git', ['-C', CLONE_ROOT, ...args], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  if (result.status !== 0) throw new Error(`git ${args.join(' ')} failed: ${result.stderr}`);
  return result.stdout;
};

const staged = git(['show', `:${PATH_IN_REPO}`]);
const broken = "reference's own presentational facts";
const fixed = "reference\\'s own presentational facts";
if (!staged.includes(broken)) {
  console.log('nothing to fix: the staged blob has no unescaped apostrophe');
  process.exit(0);
}
if (staged.includes(fixed)) {
  console.log('already escaped');
  process.exit(0);
}
const patched = staged.replace(broken, fixed);
if (patched === staged) throw new Error('replacement did not change the content');
writeFileSync(TEMP, patched);
const hash = git(['hash-object', '-w', TEMP]).trim();
git(['update-index', '--cacheinfo', `100644,${hash},${PATH_IN_REPO}`]);
console.log(`staged blob ${hash.slice(0, 12)}… with the escape fixed`);
console.log(`occurrences of the escaped form: ${(patched.match(/reference\\'s own presentational facts/gu) ?? []).length}`);

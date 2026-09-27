// Sync the DOCX working tree into the publishing clone, without touching the
// other line of work that is uncommitted there.
//
// Rule: a file is copied when it differs and is NOT among the clone's
// uncommitted paths. Generated fixtures, caches, node_modules, dist, the runtime
// and local-only files are never copied.
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, readFileSync, readdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, join, relative } from 'node:path';

const CLONE_ROOT = 'D:\\开源团队作品\\dsh-office-clone';
const WORK = 'D:\\开源团队作品\\the-last-docx';
const CLONE = join(CLONE_ROOT, 'the-last-docx');

const SKIP = /(^|[\\/])(node_modules|\.git|dist|runtime|\.build-cache|\.docx-data|_generated)([\\/]|$)|(^|[\\/])\.mcp\.json$|\.tgz$|(^|[\\/])package-lock\.json$/u;

function walk(root) {
  const files = new Set();
  const visit = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      const rel = relative(root, path).replaceAll('\\', '/');
      if (SKIP.test(rel)) continue;
      if (entry.isDirectory()) visit(path);
      else if (entry.isFile()) files.add(rel);
    }
  };
  visit(root);
  return files;
}

const sha = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');

const status = spawnSync('git', ['-C', CLONE_ROOT, 'status', '--porcelain'], { encoding: 'utf8' });
if (status.status !== 0) throw new Error(status.stderr);
const dirty = new Set(
  status.stdout.split('\n').filter(Boolean)
    .map((line) => line.slice(3).trim().replace(/^"|"$/gu, ''))
    .filter((path) => path.startsWith('the-last-docx/'))
    .map((path) => path.slice('the-last-docx/'.length)),
);
console.log(`clone has ${dirty.size} uncommitted path(s); those are never overwritten.`);

const workFiles = walk(WORK);
const cloneFiles = walk(CLONE);
const copied = [];
const added = [];
const skipped = [];

for (const rel of [...workFiles].sort()) {
  const source = join(WORK, rel);
  const target = join(CLONE, rel);
  if (!cloneFiles.has(rel)) {
    mkdirSync(dirname(target), { recursive: true });
    copyFileSync(source, target);
    added.push(rel);
    continue;
  }
  if (sha(source) === sha(target)) continue;
  if (dirty.has(rel)) { skipped.push(rel); continue; }
  copyFileSync(source, target);
  copied.push(rel);
}

console.log(`\n--- copied (updated) ${copied.length} ---`);
for (const rel of copied) console.log(`  ~ ${rel}`);
console.log(`\n--- copied (new) ${added.length} ---`);
for (const rel of added) console.log(`  + ${rel}`);
console.log(`\n--- left alone (uncommitted in the clone) ${skipped.length} ---`);
for (const rel of skipped) console.log(`  ! ${rel}`);

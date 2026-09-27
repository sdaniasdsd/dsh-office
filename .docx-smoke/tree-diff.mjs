// Compare two trees file by file: which files differ, which exist on one side
// only. Used to plan a publish without clobbering the other line of work.
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const [left, right] = process.argv.slice(2);
const SKIP = /(^|[\\/])(node_modules|\.git|dist|runtime|\.docx-data|\.codex-plugin)([\\/]|$)|\.tgz$|package-lock\.json$/u;

function walk(root) {
  const files = new Map();
  const visit = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      const rel = relative(root, path);
      if (SKIP.test(rel)) continue;
      if (entry.isDirectory()) visit(path);
      else if (entry.isFile()) files.set(rel.replaceAll('\\', '/'), createHash('sha256').update(readFileSync(path)).digest('hex'));
    }
  };
  visit(root);
  return files;
}

const a = walk(left);
const b = walk(right);
const only = (source, other) => [...source.keys()].filter((key) => !other.has(key)).sort();
const differing = [...a.keys()].filter((key) => b.has(key) && a.get(key) !== b.get(key)).sort();

console.log(`left  ${left}: ${a.size} files`);
console.log(`right ${right}: ${b.size} files`);
console.log(`\n--- only in ${right} (${only(b, a).length}) ---`);
for (const key of only(b, a)) console.log(`  + ${key}`);
console.log(`\n--- only in ${left} (${only(a, b).length}) ---`);
for (const key of only(a, b)) console.log(`  - ${key}`);
console.log(`\n--- differing (${differing.length}) ---`);
for (const key of differing) console.log(`  ~ ${key}`);

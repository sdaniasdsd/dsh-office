import { readFile, readdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { join, resolve, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// modules/<name> 是装配产物，家族目录是源码的家：这里逐文件证明两者一致。
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const overrides = new Map();
for (const raw of process.argv.slice(2)) {
  const match = /^--source=(\w+)=(.+)$/.exec(raw);
  if (match) overrides.set(match[1], match[2]);
}
const FAMILIES = [
  { family: 'docx', source: join(root, '..', 'the-last-docx'), modules: ['docx-inspect', 'docx-easy-parse', 'docx-parse', 'docx-complex-parse', 'docx-create', 'docx-styles', 'docx-edit', 'docx-render', 'docx-artifact'] },
  { family: 'pptx', source: join(root, '..', 'the-last-pptx'), modules: ['pptx-office'] },
  { family: 'xlsx', source: join(root, '..', 'the-last-xlsx'), modules: ['xlsx-office'] },
  { family: 'pdf', source: join(root, '..', 'the-last-pdf'), modules: ['pdf-office'] },
];
// 只在装配产物里存在、家族里没有对应文件的 DSH 集成文件（有意为之，逐条列出）。
const localExtensionFiles = [
  'modules/docx-edit/src/domain/docx-edit.ts',
  'modules/docx-edit/src/engine/adapter.ts',
  'modules/docx-edit/src/engine/table-xml.ts',
  'modules/docx-render/src/engine/adapter.ts',
];

const hashes = {};
const changes = [];
const familyRoots = {};
for (const entry of FAMILIES) {
  const source = resolve(overrides.get(entry.family) ?? entry.source);
  familyRoots[entry.family] = relative(root, source).split(sep).join('/');
  for (const name of entry.modules) {
    const directory = join(root, 'modules', name, 'src');
    const walk = async path => {
      for (const item of await readdir(path, { withFileTypes: true })) {
        const file = join(path, item.name);
        if (item.isDirectory()) { await walk(file); continue; }
        if (item.isFile() && !item.name.endsWith('.pyc')) {
          const key = relative(root, file).split(sep).join('/');
          const hash = createHash('sha256').update(await readFile(file)).digest('hex');
          hashes[key] = hash;
          const original = join(source, name, 'src', relative(directory, file).split(sep).join('/'));
          let sourceHash = null;
          try { sourceHash = createHash('sha256').update(await readFile(original)).digest('hex'); } catch { sourceHash = null; }
          if (sourceHash !== hash && !localExtensionFiles.includes(key)) changes.push(key);
        }
      }
    };
    try { await walk(directory); } catch { /* 模块没有 src 目录时不算差异 */ }
  }
}
await writeFile(join(root, 'modules.lock.json'), JSON.stringify({
  schema: 'dsh-office-module-provenance/v2',
  scope: 'Each modules/<name>/src tree is compared byte for byte against its family workspace; files that exist only in the assembled tree are listed in localExtensionFiles.',
  familyRoots,
  localExtensionFiles,
  allModuleSourcesMatch: changes.length === 0,
  changedSources: changes,
  sha256: hashes,
}, null, 2) + '\n');
console.log(JSON.stringify({ families: Object.keys(familyRoots).length, files: Object.keys(hashes).length, changedSources: changes }));
if (changes.length) process.exitCode = 1;

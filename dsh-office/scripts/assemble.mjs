import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// 家族目录是源码的家；modules/ 是这里复制出来的装配产物，由 provenance 逐文件比对。
// 单个家族可覆盖来源：node scripts/assemble.mjs --source docx=../../docx分区
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const overrides = new Map();
for (const raw of process.argv.slice(2)) {
  const match = /^--source=(\w+)=(.+)$/.exec(raw);
  if (match) overrides.set(match[1], match[2]);
}

const FAMILIES = [
  { family: 'docx', source: join(root, '..', 'the-last-docx'), modules: ['docx-runtime', 'docx-inspect', 'docx-easy-parse', 'docx-parse', 'docx-complex-parse', 'docx-create', 'docx-styles', 'docx-edit', 'docx-render', 'docx-artifact'] },
  { family: 'pptx', source: join(root, '..', 'the-last-pptx'), modules: ['pptx-office'] },
  { family: 'xlsx', source: join(root, '..', 'the-last-xlsx'), modules: ['xlsx-office'] },
  { family: 'pdf', source: join(root, '..', 'the-last-pdf'), modules: ['pdf-office'] },
];

const SKIP = ['node_modules', '.git', '_bench', '_generated', '__pycache__', 'coverage'];
const exists = async path => { try { await readFile(path); return true; } catch { return false; } };

let assembled = 0;
for (const entry of FAMILIES) {
  const source = resolve(overrides.get(entry.family) ?? entry.source);
  for (const name of entry.modules) {
    const from = join(source, name);
    const destination = join(root, 'modules', name);
    await mkdir(destination, { recursive: true });
    await cp(from, destination, { recursive: true, filter: path => !relative(from, path).split(sep).some(part => SKIP.includes(part)) });

    const path = join(destination, 'package.json');
    const pkg = JSON.parse(await readFile(path, 'utf8'));
    // Packaging-only alias; source capability ID and all exported behavior stay unchanged.
    if (name === 'docx-easy-parse') pkg.name = '@dsh-office-profile/docx-easy-parse';
    pkg.peerDependencies = { ...pkg.peerDependencies, 'office-core': '*', 'office-safety': '*', 'office-files': '*', 'office-test-kit': '*' };
    pkg.peerDependenciesMeta = { ...pkg.peerDependenciesMeta, ...Object.fromEntries(['office-core', 'office-safety', 'office-files', 'office-test-kit'].map(key => [key, { optional: true }])) };
    for (const field of ['dependencies', 'devDependencies']) for (const [key, value] of Object.entries(pkg[field] ?? {})) if (String(value).startsWith('file:../docx-')) pkg[field][key] = '*';
    await writeFile(path, JSON.stringify(pkg, null, 2) + '\n');

    const configPath = join(destination, 'tsconfig.json');
    if (await exists(configPath)) {
      const config = JSON.parse(await readFile(configPath, 'utf8'));
      config.include = ['src/**/*.ts', 'tests/**/*.ts', 'fixtures/**/*.ts'];
      config.exclude = ['types', 'node_modules'];
      await writeFile(configPath, JSON.stringify(config, null, 2) + '\n');
    }
    assembled += 1;
    console.log(`  ${entry.family}/${name} → modules/${name}`);
  }
}

// Promote the EXISTING canonical public type declarations into one package per owner.
// All copies keep their old type stubs for provenance, but compilation no longer includes them.
const docxSource = resolve(overrides.get('docx') ?? FAMILIES[0].source);
const shared = await readFile(join(docxSource, 'docx-inspect', 'types', 'office-deps.d.ts'), 'utf8');
for (const [, name, body] of shared.matchAll(/declare module '([^']+)' \{([\s\S]*?)\r?\n\}/g)) {
  const directory = join(root, 'packages', name);
  await mkdir(directory, { recursive: true });
  let contents = body;
  if (name === 'office-safety') contents = contents.replace(/(export interface SafetyPolicy \{)/, '$1\n    maxBlocks?: number;\n    maxTableCells?: number;');
  await writeFile(join(directory, 'index.ts'), '// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.\n' + contents + '\n');
  await writeFile(join(directory, 'package.json'), JSON.stringify({ name, version: '0.1.0', private: true, type: 'module', types: 'index.ts', exports: { '.': './index.ts' } }, null, 2) + '\n');
}
console.log(`Assembled ${assembled} modules from ${FAMILIES.length} families into ${root}. No source module was modified.`);

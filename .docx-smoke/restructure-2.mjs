// 阶段 2：写四个家族的清单/README，并把外壳里的脚本与清单改成从「家族」组装。
//
// 设计：家族目录是**源码的家**；外壳里的 modules/ 是 assemble 从四个家族复制出来的**装配产物**，
// 由 provenance 逐文件比对保证一致。外壳内部的布局（modules/<name>/src/...）保持不变，
// 因此构建脚本、引擎脚本拷贝、pnpm workspace 链接都不用改。
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'D:\\开源团队作品\\dsh-office-clone';
const SHELL = join(REPO, 'dsh-office');
const write = (path, contents) => { writeFileSync(path, contents, 'utf8'); console.log(`  wrote ${path.replace(REPO + '\\', '')}`); };

const FAMILIES = [
  {
    dir: 'the-last-docx',
    name: '@dsh-office-profile/docx-workspace',
    glob: 'docx-*',
    readme: `# the-last-docx — DOCX 家族工作区

九个 DOCX 模块的**源码家目录**（每个模块一个包，包名 \`@dsh-office-profile/docx-*\`）：

| 模块 | 做什么 |
| --- | --- |
| \`docx-inspect\` | 输入探测与安全读取（OOXML 结构、宏、外链、嵌入对象） |
| \`docx-easy-parse\` | 保留旧版输出形态的简化解析（legacy IR） |
| \`docx-parse\` | 结构化解析：语义视图 + 物理视图 + 源映射（dual IR） |
| \`docx-complex-parse\` | 复杂版面：逻辑表格栅格、分页关系、浮动对象、置信度 |
| \`docx-create\` | 按计划新建文档（场景 → 装帧寄存器、表格、编号、页面几何） |
| \`docx-styles\` | styles.xml 的命名样式编排与解析 |
| \`docx-edit\` | 就地编辑（修订、批注、段落/表格格式），未触及部分字节不变 |
| \`docx-render\` | 渲染成 PDF 与逐页图片，供视觉复核 |
| \`docx-artifact\` | 版本化交付与不可变清单 |

## 怎么用

这个目录只放源码；依赖与链接由外壳统一管理：

\`\`\`bash
cd ../dsh-office
pnpm install          # 把 @dsh-office-profile/* 链接进 node_modules
pnpm assemble         # 把四个家族复制进 dsh-office/modules/（装配产物）
pnpm provenance       # 逐文件校验 modules/ 与家族源码一致
pnpm test:modules     # 逐模块 typecheck + 测试
\`\`\`

> 历史：这九个模块原在本机仓库外的 \`docx分区\` 里开发，再由 \`assemble\` 复制进仓库。
> 现在仓库内的本目录就是它们的家；仍想在外部目录开发时用
> \`pnpm assemble --source docx=<外部路径>\`（外壳里也留了 \`assemble:docx-partition\` 脚本）。
`,
  },
  {
    dir: 'the-last-pptx',
    name: '@dsh-office-profile/pptx-workspace',
    glob: 'pptx-*',
    readme: `# the-last-pptx — PPTX 家族工作区

PPTX 相关模块的**源码家目录**。

| 模块 | 做什么 |
| --- | --- |
| \`pptx-office\` | 幻灯片检查与抽取、按坐标替换文本、按标题/正文层级做受控排版；**不新建 deck**（建 deck 用 python-pptx，本模块随后做排版与校验） |

## 怎么用

\`\`\`bash
cd ../dsh-office
pnpm install && pnpm assemble && pnpm provenance && pnpm test:modules
\`\`\`

运行时不另装：随插件分发的 Python 里带 \`python-pptx\`（\`dsh-office/runtime/win32-x64/python\`），
引擎脚本是 \`pptx-office/src/engine/pptx_bridge.py\`，由 \`scripts/build-dsh.mjs\` 拷进包内。
`,
  },
  {
    dir: 'the-last-xlsx',
    name: '@dsh-office-profile/xlsx-workspace',
    glob: 'xlsx-*',
    readme: `# the-last-xlsx — XLSX（表格）家族工作区

表格相关模块的**源码家目录**。

| 模块 | 做什么 |
| --- | --- |
| \`xlsx-office\` | 工作簿检查与结构抽取、读写单元格、打印版式安全修复；宏文件不支持，公式不重算 |

## 怎么用

\`\`\`bash
cd ../dsh-office
pnpm install && pnpm assemble && pnpm provenance && pnpm test:modules
\`\`\`

依赖 \`exceljs\`、\`@xmldom/xmldom\`、\`fflate\`，由外壳的 workspace 统一安装。
`,
  },
];

for (const family of FAMILIES) {
  write(join(REPO, family.dir, 'package.json'), `${JSON.stringify({
    name: family.name, version: '0.1.0', private: true, type: 'module',
    workspaces: [family.glob],
    scripts: { assemble: 'node ../dsh-office/scripts/assemble.mjs', 'test:modules': 'node ../dsh-office/scripts/check-modules.mjs' },
  }, null, 2)}\n`);
  write(join(REPO, family.dir, 'README.md'), family.readme);
}

// the-last-pdf 已有自己的 workspace 清单，只补一节说明 pdf-office 的位置
const pdfReadmePath = join(REPO, 'the-last-pdf', 'README.md');
const pdfReadme = readFileSync(pdfReadmePath, 'utf8');
if (!pdfReadme.includes('pdf-office')) {
  write(pdfReadmePath, `${pdfReadme.trimEnd()}

## DSH 集成模块

\`pdf-office/\` 是插件里被注册的那个 PDF 模块（\`@dsh-office-profile/pdf-office\`）：只读检查与文本抽取，
基于 \`pdfjs-dist\`。其余 \`pdf-*\` 目录是本家族自己的实现与文档工作区；两者都由 \`../dsh-office\` 统一装配与打包。
`);
}

// 外壳清单改名
const shellPackagePath = join(SHELL, 'package.json');
const shellPackage = JSON.parse(readFileSync(shellPackagePath, 'utf8'));
shellPackage.name = 'dsh-office';
shellPackage.description = 'DSH Office 插件外壳：DOCX / PDF / PPTX / XLSX 四个家族模块的装配、注册、打包与交付。';
shellPackage.scripts['assemble:docx-partition'] = 'node scripts/assemble.mjs --source docx=../../docx分区';
write(shellPackagePath, `${JSON.stringify(shellPackage, null, 2)}\n`);

// ---------------------------------------------------------------------------
// 新的 assemble.mjs：从四个家族组装，可对单个家族覆盖来源
// ---------------------------------------------------------------------------
const assemble = `import { cp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// 家族目录是源码的家；modules/ 是这里复制出来的装配产物，由 provenance 逐文件比对。
// 单个家族可覆盖来源：node scripts/assemble.mjs --source docx=../../docx分区
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const overrides = new Map();
for (const raw of process.argv.slice(2)) {
  const match = /^--source=(\\w+)=(.+)$/.exec(raw);
  if (match) overrides.set(match[1], match[2]);
}

const FAMILIES = [
  { family: 'docx', source: join(root, '..', 'the-last-docx'), modules: ['docx-inspect', 'docx-easy-parse', 'docx-parse', 'docx-complex-parse', 'docx-create', 'docx-styles', 'docx-edit', 'docx-render', 'docx-artifact'] },
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
    await writeFile(path, JSON.stringify(pkg, null, 2) + '\\n');

    const configPath = join(destination, 'tsconfig.json');
    if (await exists(configPath)) {
      const config = JSON.parse(await readFile(configPath, 'utf8'));
      config.include = ['src/**/*.ts', 'tests/**/*.ts', 'fixtures/**/*.ts'];
      config.exclude = ['types', 'node_modules'];
      await writeFile(configPath, JSON.stringify(config, null, 2) + '\\n');
    }
    assembled += 1;
    console.log(\`  \${entry.family}/\${name} → modules/\${name}\`);
  }
}

// Promote the EXISTING canonical public type declarations into one package per owner.
// All copies keep their old type stubs for provenance, but compilation no longer includes them.
const docxSource = resolve(overrides.get('docx') ?? FAMILIES[0].source);
const shared = await readFile(join(docxSource, 'docx-inspect', 'types', 'office-deps.d.ts'), 'utf8');
for (const [, name, body] of shared.matchAll(/declare module '([^']+)' \\{([\\s\\S]*?)\\r?\\n\\}/g)) {
  const directory = join(root, 'packages', name);
  await mkdir(directory, { recursive: true });
  let contents = body;
  if (name === 'office-safety') contents = contents.replace(/(export interface SafetyPolicy \\{)/, '$1\\n    maxBlocks?: number;\\n    maxTableCells?: number;');
  await writeFile(join(directory, 'index.ts'), '// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.\\n' + contents + '\\n');
  await writeFile(join(directory, 'package.json'), JSON.stringify({ name, version: '0.1.0', private: true, type: 'module', types: 'index.ts', exports: { '.': './index.ts' } }, null, 2) + '\\n');
}
console.log(\`Assembled \${assembled} modules from \${FAMILIES.length} families into \${root}. No source module was modified.\`);
`;
write(join(SHELL, 'scripts', 'assemble.mjs'), assemble);

// ---------------------------------------------------------------------------
// 新的 provenance.mjs：modules/ 与四个家族逐文件比对
// ---------------------------------------------------------------------------
const provenance = `import { readFile, readdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { join, resolve, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// modules/<name> 是装配产物，家族目录是源码的家：这里逐文件证明两者一致。
const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
const overrides = new Map();
for (const raw of process.argv.slice(2)) {
  const match = /^--source=(\\w+)=(.+)$/.exec(raw);
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
}, null, 2) + '\\n');
console.log(JSON.stringify({ families: Object.keys(familyRoots).length, files: Object.keys(hashes).length, changedSources: changes }));
if (changes.length) process.exitCode = 1;
`;
write(join(SHELL, 'scripts', 'provenance.mjs'), provenance);

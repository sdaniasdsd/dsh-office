// 阶段 1：把仓库整理成「四个家族 + 一个外壳」的物理布局。
//
// 用文件系统改名而不是 `git mv`：外壳目录里有 node_modules / dist / runtime 这些被忽略的重目录，
// git mv 只会搬 tracked 文件、把它们留在原地。改名后由 git 自己做重命名识别（内容相同即识别为 rename）。
//
// 顺序很重要：先把外壳改名（腾出 the-last-docx 这个名字），再把各家族模块从 modules/ 里搬出去。
import { existsSync, mkdirSync, readdirSync, renameSync, statSync } from 'node:fs';
import { join } from 'node:path';

const REPO = 'D:\\开源团队作品\\dsh-office-clone';
const SHELL_OLD = join(REPO, 'the-last-docx');
const SHELL_NEW = join(REPO, 'dsh-office');
const DOCX_FAMILY = join(REPO, 'the-last-docx');
const PPTX_FAMILY = join(REPO, 'the-last-pptx');
const XLSX_FAMILY = join(REPO, 'the-last-xlsx');
const PDF_FAMILY = join(REPO, 'the-last-pdf');

const DOCX_MODULES = ['docx-inspect', 'docx-easy-parse', 'docx-parse', 'docx-complex-parse', 'docx-create', 'docx-styles', 'docx-edit', 'docx-render', 'docx-artifact'];

const move = (from, to) => {
  if (!existsSync(from)) throw new Error(`missing: ${from}`);
  mkdirSync(join(to, '..'), { recursive: true });
  if (existsSync(to)) throw new Error(`target already exists: ${to}`);
  renameSync(from, to);
  console.log(`  moved ${from.replace(REPO + '\\', '')}  →  ${to.replace(REPO + '\\', '')}`);
};

console.log('1) 外壳改名 the-last-docx → dsh-office');
move(SHELL_OLD, SHELL_NEW);

console.log('\n2) 建家族目录');
for (const family of [DOCX_FAMILY, PPTX_FAMILY, XLSX_FAMILY]) {
  mkdirSync(family, { recursive: true });
  console.log(`  created ${family.replace(REPO + '\\', '')}`);
}

console.log('\n3) 各家族模块搬出 modules/');
const modulesDir = join(SHELL_NEW, 'modules');
for (const name of DOCX_MODULES) move(join(modulesDir, name), join(DOCX_FAMILY, name));
move(join(modulesDir, 'pptx-office'), join(PPTX_FAMILY, 'pptx-office'));
move(join(modulesDir, 'xlsx-office'), join(XLSX_FAMILY, 'xlsx-office'));
move(join(modulesDir, 'pdf-office'), join(PDF_FAMILY, 'pdf-office'));

console.log(`\nremaining in dsh-office/modules: ${readdirSync(modulesDir).length ? readdirSync(modulesDir).join(', ') : '(empty)'}`);
for (const [label, dir] of [['the-last-docx', DOCX_FAMILY], ['the-last-pptx', PPTX_FAMILY], ['the-last-xlsx', XLSX_FAMILY], ['the-last-pdf', PDF_FAMILY]]) {
  const entries = readdirSync(dir, { withFileTypes: true }).filter((entry) => entry.isDirectory()).map((entry) => entry.name);
  console.log(`${label}: ${entries.length} 个目录 → ${entries.slice(0, 12).join(', ')}${entries.length > 12 ? ' …' : ''}`);
}
console.log(`\nshell still has: ${readdirSync(SHELL_NEW, { withFileTypes: true }).filter((entry) => entry.isDirectory()).map((entry) => entry.name).join(', ')}`);
console.log(`dist present: ${existsSync(join(SHELL_NEW, 'dist', 'dsh-docx'))}  node_modules present: ${statSync(join(SHELL_NEW, 'node_modules')).isDirectory()}`);

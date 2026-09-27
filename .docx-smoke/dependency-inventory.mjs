// 依赖清点：把五个顶层目录里的清单、Python 需求文件、以及引擎脚本真正 import 的第三方包都读出来。
// 目的不是"猜"，而是把每条依赖的**声明位置**找出来，包括那些没被声明的。
import { readdirSync, readFileSync, existsSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const REPO = 'D:\\开源团队作品\\dsh-office-clone';
const ROOTS = ['the-last-docx', 'the-last-pdf', 'the-last-pptx', 'the-last-xlsx', 'dsh-office'];
const SKIP = /(^|[\\/])(node_modules|\.git|dist|\.build-cache|_generated|__pycache__|\.vite)([\\/]|$)/u;

const walk = (root, match, out = []) => {
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    const path = join(root, entry.name);
    const rel = relative(REPO, path);
    if (SKIP.test(rel)) continue;
    if (entry.isDirectory()) walk(path, match, out);
    else if (match(entry.name)) out.push(path);
  }
  return out;
};

const manifests = [];
for (const root of ROOTS) for (const path of walk(join(REPO, root), (name) => name === 'package.json')) manifests.push(path);

console.log('=== package.json（按顶层目录） ===');
const nodeDeps = new Map();
for (const path of manifests.sort()) {
  const pkg = JSON.parse(readFileSync(path, 'utf8'));
  const rel = relative(REPO, path).replaceAll('\\', '/');
  const groups = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'].filter((field) => pkg[field] && Object.keys(pkg[field]).length);
  console.log(`\n${rel}   name=${pkg.name}  private=${pkg.private ?? false}`);
  for (const field of groups) {
    const entries = Object.entries(pkg[field]);
    console.log(`  ${field}: ${entries.map(([key, value]) => `${key}@${value}`).join(', ')}`);
    if (field === 'dependencies' || field === 'devDependencies') {
      for (const [key, value] of entries) {
        if (String(value).startsWith('file:') || String(value) === '*') continue;
        const key2 = `${key}@${value}`;
        const where = nodeDeps.get(key2) ?? [];
        where.push(rel);
        nodeDeps.set(key2, where);
      }
    }
  }
}

console.log('\n=== 汇总：npm 依赖（谁声明的） ===');
for (const [key, where] of [...nodeDeps].sort()) console.log(`  ${key.padEnd(48)} ${where.length} 处：${where.slice(0, 3).join(', ')}${where.length > 3 ? ' …' : ''}`);

console.log('\n=== Python 需求文件 ===');
for (const root of ROOTS) for (const path of walk(join(REPO, root), (name) => /^requirements.*\.txt$/u.test(name))) {
  console.log(`\n${relative(REPO, path).replaceAll('\\', '/')}`);
  console.log(readFileSync(path, 'utf8').trim().split('\n').map((line) => `  ${line}`).join('\n'));
}

console.log('\n=== 引擎脚本里 import 的第三方模块 ===');
const STDLIB = new Set(['sys', 'os', 'json', 're', 'zipfile', 'hashlib', 'pathlib', 'argparse', 'io', 'tempfile', 'shutil', 'subprocess', 'math', 'itertools', 'collections', 'typing', 'dataclasses', 'xml', 'contextlib', 'importlib', 'unicodedata', 'struct', 'base64', 'time', 'warnings', 'textwrap', 'functools', 'glob', 'csv', 'decimal', 'random', 'html', 'uuid', 'statistics', 'traceback', 'copy', 'string', 'binascii', 'zipfile', 'lzma', 'zlib', 'ctypes', 'errno', 'platform', 'posixpath', 'ntpath', 'shlex', 'sqlite3', 'threading', 'queue', 'signal', 'socket', 'ssl', 'urllib', 'http', 'email', 'secrets', 'datetime', 'operator', 'abc', 'enum', 'metaclass', 'inspect', 'ast', 'tokenize', 'token', 'dis', 'pickle', 'shelve', 'dbm', 'weakref', 'gc', 'sysconfig', 'site', 'venv', 'ensurepip', 'runpy', 'pkgutil', 'modulefinder', 'pprint', 'reprlib', 'graphlib', 'numbers', 'cmath', 'fractions', 'array', 'bisect', 'heapq', 'keyword', 'difflib', 'codecs', 'locale', 'gettext', 'logging', 'getpass', 'curses', 'readline', 'rlcompleter', 'pydoc', 'doctest', 'unittest', 'argparse']);
const imports = new Map();
for (const root of ROOTS) for (const path of walk(join(REPO, root), (name) => name.endsWith('.py'))) {
  const source = readFileSync(path, 'utf8');
  for (const match of source.matchAll(/^\s*(?:from|import)\s+([A-Za-z_][\w.]*)/gmu)) {
    const top = match[1].split('.')[0];
    if (STDLIB.has(top)) continue;
    const list = imports.get(top) ?? [];
    list.push(relative(REPO, path).replaceAll('\\', '/'));
    imports.set(top, list);
  }
}
for (const [name, where] of [...imports].sort()) console.log(`  ${name.padEnd(20)} ${where.length} 个文件，例如 ${where[0]}`);

console.log('\n=== 外部运行时（随包分发的那些） ===');
const runtime = join(REPO, 'dsh-office', 'runtime');
console.log(`  dsh-office/runtime 存在: ${existsSync(runtime)}  内容: ${existsSync(runtime) ? readdirSync(runtime).join(', ') : '(无)'}`);
for (const probe of ['dist/dsh-docx', 'scripts/fetch-dsh-runtime.ps1']) {
  console.log(`  ${probe}: ${existsSync(join(REPO, 'dsh-office', probe))}`);
}

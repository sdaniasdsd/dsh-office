// 随包分发的 Python 里到底装了哪些第三方包：逐个 import，缺哪个记哪个（一次 import 失败不影响后面的）。
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

const PYTHON = process.argv[2] ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64', 'python', 'python.exe');
const MODULES = [
  'lxml', 'pptx', 'PIL', 'oletools', 'docling', 'rdocx', 'openpyxl', 'xlsxwriter', 'XlsxWriter',
  'typing_extensions', 'pdfjs', 'pypdf', 'fitz', 'numpy', 'pandas', 'six', 'et_xmlfile', 'olefile',
];

console.log(`python: ${PYTHON}`);
const version = spawnSync(PYTHON, ['-c', 'import sys; print(sys.version.split()[0])'], { encoding: 'utf8', windowsHide: true });
console.log(`version: ${(version.stdout ?? '').trim()}`);

for (const name of MODULES) {
  const probe = spawnSync(PYTHON, ['-c', `import ${name} as m; print(getattr(m, '__version__', '?'))`], { encoding: 'utf8', windowsHide: true });
  const ok = probe.status === 0;
  console.log(`  ${name.padEnd(20)} ${ok ? `present  ${(probe.stdout ?? '').trim()}` : 'MISSING'}`);
}

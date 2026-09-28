// 绕开 MCP 的 JSON 边界，直接调构建产物导出的 doctor()，把「不可丢失序列化」的值找出来。
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const REPO = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office';
const DIST = join(REPO, 'dist', 'dsh-docx');
const core = await import(pathToFileURL(join(DIST, 'lib', 'index.mjs')).href);

const options = {
  pythonPath: join(REPO, 'runtime', 'win32-x64', 'python', 'python.exe'),
  sofficePath: join(REPO, 'runtime', 'win32-x64', 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(REPO, 'runtime', 'win32-x64', 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
  runtimeRoot: join(REPO, 'runtime'),
};
const report = await core.doctor(options);

// 复刻 docx-artifact 的 jsonValue 判据，只报告问题字段
const problems = [];
const walk = (value, path, seen = new Set()) => {
  if (value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value))) return;
  if (typeof value === 'undefined') { problems.push(`${path} = undefined`); return; }
  if (typeof value !== 'object') { problems.push(`${path} = ${typeof value}`); return; }
  if (seen.has(value)) { problems.push(`${path} = 重复引用/环`); return; }
  const proto = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && proto !== Object.prototype && proto !== null) { problems.push(`${path} = 非普通对象 (${proto?.constructor?.name ?? 'null proto'})`); return; }
  seen.add(value);
  for (const [key, child] of Object.entries(value)) walk(child, `${path}.${key}`, seen);
  seen.delete(value);
};
walk(report, 'report');
console.log(`doctor() 顶层键: ${Object.keys(report).join(', ')}`);
console.log(`runtimeResolution = ${JSON.stringify(report.runtimeResolution, null, 1).slice(0, 600)}`);
console.log(`\n会被 JSON 边界拒绝的值（${problems.length} 个）:`);
for (const line of problems) console.log(`  ${line}`);

// 用**构建产物**（dist/dsh-docx）起服务，看 doctor 现在怎么报 exceljs / xlsx。
// 运行时用仓库里的 runtime/win32-x64（刚从 dsh-toolchain 取回来的那份）。
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const SHELL = 'D:\\开源团队作品\\dsh-office-clone\\dsh-office';
const CORE = join(SHELL, 'dist', 'dsh-docx');
const RUNTIME = join(SHELL, 'runtime', 'win32-x64');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';

for (const [label, path] of [['core', join(CORE, 'lib', 'server.mjs')], ['bundled.json', join(CORE, 'lib', 'bundled.json')], ['runtime', join(RUNTIME, 'python', 'python.exe')]]) {
  console.log(`${label.padEnd(13)} ${existsSync(path) ? 'ok' : '缺失'}  ${path}`);
}

const child = spawn(process.execPath, [join(CORE, 'lib', 'server.mjs')], {
  env: {
    ...process.env,
    THE_LAST_DOCX_WORKSPACE: WORKSPACE,
    THE_LAST_DOCX_DATA: join(SHELL, 'dist', 'doctor-probe-data'),
    DOCX_PYTHON: join(RUNTIME, 'python', 'python.exe'),
    DOCX_SOFFICE: join(RUNTIME, 'libreoffice', 'program', 'soffice.com'),
    DOCX_PDFTOPPM: join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
  },
  stdio: ['pipe', 'pipe', 'pipe'],
});
let buffer = ''; let id = 1; const pending = new Map();
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk;
  let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, index).trim(); buffer = buffer.slice(index + 1);
    if (!line) continue;
    let message; try { message = JSON.parse(line); } catch { continue; }
    if (message.id !== undefined && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  }
});
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => process.stderr.write(chunk));
const request = (method, params) => { const current = id++; const promise = new Promise((resolve) => pending.set(current, resolve)); child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: current, method, params })}\n`); return promise; };
await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'doctor-built', version: '1' } });
child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })}\n`);
const reply = await request('tools/call', { name: 'docx_doctor', arguments: {} });
const report = JSON.parse(reply.result.content.find((part) => part.type === 'text').text);
console.log(`\ncapabilities : ${JSON.stringify(report.capabilities)}`);
console.log(`exceljs      : ${JSON.stringify(report.exceljs)}`);
console.log(`pdfjs        : ${JSON.stringify(report.pdfjs)}`);
console.log(`libreoffice  : ${JSON.stringify(report.libreoffice)}`);
console.log(`bundled 包数 : ${(report.bundledPackages ?? []).length}（含 exceljs=${(report.bundledPackages ?? []).includes('exceljs')}）`);
child.kill();
process.exit(report.capabilities.xlsx ? 0 : 1);

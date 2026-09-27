// 用已安装的 DSH Office 插件生成演讲稿：建文档 → 渲染 → 回读校验 → 落盘。
//
// 为什么直接驱动插件自己的 server.mjs，而不是走本会话里的 mcp__docx__ 工具：本会话里那个
// 进程启动于 09/26 20:58，内存里是旧构建（它的 docx-parse schema 里没有 parseFormatting）。
// 这里拉起的是刚装好的 0.8.1，用的是它自带的运行时。
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { SPEECH_DOCUMENT } from './speech-content.mjs';

const PACKAGE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
const SERVER = join(PACKAGE, 'lib', 'server.mjs');
const RUNTIME = join(PACKAGE, 'runtime', 'win32-x64');
const WORKSPACE = 'C:\\Users\\AA\\Desktop\\bench';
const DATA = 'D:\\认真版agent\\.docx-smoke\\data-speech';
const OUT_DIR = 'C:\\Users\\AA\\Desktop\\bench\\out-speech';
const TMP = join(DATA, '_tmp');

for (const path of [DATA, TMP, OUT_DIR]) mkdirSync(path, { recursive: true });

const child = spawn(process.execPath, [SERVER], {
  env: {
    ...process.env,
    THE_LAST_DOCX_WORKSPACE: WORKSPACE,
    THE_LAST_DOCX_DATA: DATA,
    DOCX_PYTHON: join(RUNTIME, 'python', 'python.exe'),
    DOCX_SOFFICE: join(RUNTIME, 'libreoffice', 'program', 'soffice.com'),
    DOCX_PDFTOPPM: join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
    TEMP: TMP, TMP,
    PATH: `${join(RUNTIME, 'libreoffice', 'System64')};${process.env.PATH ?? ''}`,
  },
  stdio: ['pipe', 'pipe', 'pipe'],
});

let nextId = 1;
const pending = new Map();
let buffer = '';
child.stdout.setEncoding('utf8');
child.stdout.on('data', (chunk) => {
  buffer += chunk;
  let index;
  while ((index = buffer.indexOf('\n')) >= 0) {
    const line = buffer.slice(0, index).trim();
    buffer = buffer.slice(index + 1);
    if (!line) continue;
    let message;
    try { message = JSON.parse(line); } catch { continue; }
    if (message.id !== undefined && pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  }
});
const stderr = [];
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => { stderr.push(chunk); process.stderr.write(chunk); });
child.on('exit', (code) => { if (code !== 0) process.stderr.write(`\n[server exited: ${code}]\n`); });

const request = (method, params) => {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
};
const notify = (method, params) => child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
function payload(response) {
  const result = response.result;
  if (!result) return { error: response.error ?? 'no result' };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
}
const callTool = async (name, args) => payload(await request('tools/call', { name, arguments: args }));
async function readSpilled(ref) {
  let text = ''; let offset = 0;
  for (;;) {
    const chunk = await callTool('docx_read_artifact', { artifactRef: ref, offset, length: 200000 });
    if (chunk.error) throw new Error(`spill read failed: ${chunk.error}`);
    text += chunk.value.text ?? '';
    if (chunk.value.complete) break;
    offset = chunk.value.nextOffset;
  }
  return JSON.parse(text);
}
const localPath = (uri) => decodeURIComponent(String(uri).replace(/^file:\/\/\//u, '').replace(/^file:\/\//u, '').replace(/^file:/u, ''));
const sha256 = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'speech-run', version: '1' } });
notify('notifications/initialized');
const listed = await request('tools/list', {});
const tools = (listed.result?.tools ?? []).map((tool) => tool.name);
console.log(`serverInfo : ${JSON.stringify(initialize.result?.serverInfo)}`);
console.log(`tools      : ${tools.join(', ')}\n`);

// 1. 生成文档
const created = await callTool('docx_call', {
  moduleId: 'docx-create', operation: 'execute',
  input: { requestId: 'speech-create', plan: { kind: 'create', document: SPEECH_DOCUMENT } },
});
if (created.error) throw new Error(`create failed: ${created.error}`);
const createdResult = created.value.result ?? await readSpilled(created.value.resultRef);
const doc = createdResult.artifactRef;
console.log(`created    : ${doc.sha256?.slice(0, 12)}… ${doc.sizeBytes ?? ''}B  register=${createdResult.design?.register} decoration=${createdResult.design?.decoration}`);
console.log(`verify     : ${JSON.stringify(createdResult.verification ?? null)}`);

// 2. 回读校验：段落数、字号、缩进
const parsed = await callTool('docx_call', {
  moduleId: 'docx-parse', operation: 'execute',
  input: { artifactRef: doc, requestId: 'speech-parse', options: { featureFlags: { parseFormatting: true } } },
});
if (parsed.error) throw new Error(`parse failed: ${parsed.error}`);
const parsedResult = parsed.value.result ?? await readSpilled(parsed.value.resultRef);
const ir = parsedResult.result?.ir ?? parsedResult.ir;
const blocks = ir?.content?.semantic?.blocks ?? [];
const kinds = blocks.reduce((counts, block) => ({ ...counts, [block.kind]: (counts[block.kind] ?? 0) + 1 }), {});
const formatting = ir?.content?.formatting;
const bySize = new Map();
for (const entry of formatting?.paragraphs ?? []) {
  for (const run of entry.runs) if (run.size !== undefined && run.text.trim()) bySize.set(run.size, (bySize.get(run.size) ?? 0) + 1);
}
const indents = new Map();
for (const entry of formatting?.paragraphs ?? []) {
  const key = entry.indent === undefined ? '(none)' : JSON.stringify(entry.indent);
  indents.set(key, (indents.get(key) ?? 0) + 1);
}
const characters = blocks.filter((block) => block.kind === 'paragraph').reduce((total, block) => total + String(block.text ?? '').length, 0);
console.log(`parsed     : blocks=${blocks.length} ${JSON.stringify(kinds)}  正文字数=${characters}`);
console.log(`sizes      : ${[...bySize].sort((a, b) => b[0] - a[0]).map(([size, count]) => `${size}pt×${count}`).join('  ')}`);
console.log(`indents    : ${[...indents].map(([key, count]) => `${key}×${count}`).join('  ')}`);

// 3. 渲染成 PDF 与页面图
const rendered = await callTool('docx_call', {
  moduleId: 'docx-render', operation: 'execute',
  input: { artifactRef: doc, requestId: 'speech-render', config: { limits: { dpi: 144 } } },
});
if (rendered.error) throw new Error(`render failed: ${rendered.error}`);
const renderResult = rendered.value.result ?? await readSpilled(rendered.value.resultRef);
console.log(`rendered   : ${renderResult.pageCount} pages, pdf ${renderResult.pdf?.sizeBytes ?? '?'}B, ${(renderResult.pages ?? []).length} page image(s)`);

// 4. 落盘
const deliverable = join(OUT_DIR, '嵌入式的发展路径-演讲稿.docx');
copyFileSync(localPath(doc.uri), deliverable);
const pdf = join(OUT_DIR, '嵌入式的发展路径-演讲稿.pdf');
copyFileSync(localPath(renderResult.pdf.uri), pdf);
const images = [];
for (const page of renderResult.pages ?? []) {
  const target = join(OUT_DIR, `page-${page.index ?? images.length + 1}.png`);
  copyFileSync(localPath(page.image.uri), target);
  images.push(target);
}
console.log(`\ndocx       : ${deliverable}\n  sha256   : ${sha256(deliverable)}  ${readFileSync(deliverable).length}B`);
console.log(`pdf        : ${pdf}\n  sha256   : ${sha256(pdf)}  ${readFileSync(pdf).length}B`);
for (const image of images) console.log(`png        : ${image}  ${readFileSync(image).length}B`);

writeFileSync(join(OUT_DIR, 'run-summary.json'), JSON.stringify({
  server: { path: SERVER, version: JSON.parse(readFileSync(join(PACKAGE, 'package.json'), 'utf8')).version, serverInfo: initialize.result?.serverInfo },
  tools,
  document: { path: deliverable, sha256: sha256(deliverable), sizeBytes: readFileSync(deliverable).length },
  pdf: { path: pdf, sha256: sha256(pdf), sizeBytes: readFileSync(pdf).length },
  pages: renderResult.pageCount,
  design: createdResult.design ?? null,
  verification: createdResult.verification ?? null,
  blocks: kinds,
  characters,
  sizes: Object.fromEntries(bySize),
  indents: Object.fromEntries(indents),
}, null, 2));

child.kill();
console.log(stderr.length ? '\n(server wrote to stderr above)' : '');
process.exit(0);

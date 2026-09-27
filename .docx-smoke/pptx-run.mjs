// 把初稿交给插件处理：inspect → extract → formatText（统一标题/正文字号层级）→ verify，
// 再渲染成 PDF 与逐页 PNG 落盘。
//
// 说明：插件本身没有「新建 deck」的能力（pptx-office 只有 inspect / extract /
// replaceText / formatText），所以初稿由 python-pptx 生成，再由插件做排版层级与校验。
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const PACKAGE = join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx');
const RUNTIME = join(PACKAGE, 'runtime', 'win32-x64');
const PYTHON = join(RUNTIME, 'python', 'python.exe');
const SOFFICE = join(RUNTIME, 'libreoffice', 'program', 'soffice.com');
const PDFTOPPM = join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe');
const SCRATCH = 'D:\\认真版agent\\.docx-smoke';
const DATA = join(SCRATCH, 'data-pptx');
const DRAFT = join(DATA, '嵌入式的发展路径-配套PPT(初稿).pptx');
const OUT_DIR = 'C:\\Users\\AA\\Desktop\\bench\\out-speech';
const TMP = join(DATA, '_tmp');
for (const path of [DATA, TMP, OUT_DIR]) mkdirSync(path, { recursive: true });

const child = spawn(process.execPath, [join(PACKAGE, 'lib', 'server.mjs')], {
  env: {
    ...process.env,
    THE_LAST_DOCX_WORKSPACE: SCRATCH,
    THE_LAST_DOCX_DATA: DATA,
    DOCX_PYTHON: PYTHON,
    DOCX_SOFFICE: SOFFICE,
    DOCX_PDFTOPPM: PDFTOPPM,
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
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => process.stderr.write(chunk));
const request = (method, params) => {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
};
const notify = (method, params) => child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
const callTool = async (name, args) => {
  const response = await request('tools/call', { name, arguments: args });
  const result = response.result;
  if (!result) return { error: response.error ?? 'no result' };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
};
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
const unwrap = async (value) => (value?.resultRef ? readSpilled(value.resultRef) : (value?.result ?? value));
const localPath = (uri) => decodeURIComponent(String(uri).replace(/^file:\/\/\//u, '').replace(/^file:\/\//u, '').replace(/^file:/u, ''));
const sha256 = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'pptx-run', version: '1' } });
notify('notifications/initialized');
console.log(`serverInfo : ${JSON.stringify(initialize.result?.serverInfo)}`);

// 1. 导入初稿
const imported = await callTool('docx_import', { path: DRAFT });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const draft = imported.value;
console.log(`draft      : ${draft.sha256.slice(0, 12)}… ${draft.sizeBytes}B  ${draft.label}`);

// 2. inspect
const inspected = await callTool('pptx_call', { operation: 'inspect', input: { requestId: 'pptx-inspect', operation: 'inspect', artifactRef: draft } });
if (inspected.error) throw new Error(`inspect failed: ${inspected.error}`);
const inspection = await unwrap(inspected.value);
console.log(`inspect    : slides=${inspection.slideCount} shapes=${inspection.shapeCount} ${inspection.widthEmu}×${inspection.heightEmu} EMU  engine=${inspection.engine}`);

// 3. extract（拿到插件认可的 shape/段落/run 坐标）
const extracted = await callTool('pptx_call', { operation: 'execute', input: { requestId: 'pptx-extract', operation: 'execute', artifactRef: draft, payload: { action: 'extract' } } });
if (extracted.error) throw new Error(`extract failed: ${extracted.error}`);
const extraction = await unwrap(extracted.value);
const slides = extraction.slides ?? [];
console.log(`extract    : ${slides.length} slides, ${slides.reduce((total, slide) => total + (slide.shapes?.length ?? 0), 0)} shapes`);
for (const slide of slides.slice(0, 2)) {
  const titles = (slide.shapes ?? []).map((shape) => `${shape.shapeId}:${shape.kind}/${String(shape.text ?? '').slice(0, 18)}`);
  console.log(`  slide ${slide.slideNumber}: ${titles.join('  ')}`);
}

// 4. formatText：交给插件统一标题/正文字号层级
const formatted = await callTool('pptx_call', {
  operation: 'execute',
  input: {
    requestId: 'pptx-formatText', operation: 'execute', artifactRef: draft,
    payload: { action: 'formatText', changes: [{ scope: 'allSlides', titleFontSize: 30, bodyFontSize: 18, accentColor: '#1F4E79' }] },
  },
});
if (formatted.error) throw new Error(`formatText failed: ${formatted.error}`);
const formatting = await unwrap(formatted.value);
const formattedRef = formatting.artifactRef;
console.log(`formatText : formattedRuns=${formatting.formattedRuns} formattedSlides=${formatting.formattedSlides} verified=${formatting.verifiedFormattingRuns} contrastAdjusted=${JSON.stringify(formatting.contrastAdjustedSlides)} titleColorPreserved=${JSON.stringify(formatting.titleColorPreservedSlides)}`);
console.log(`           : visualReview=${formatting.visualReview}`);

// 5. verify 插件自己的检查
const verified = await callTool('pptx_call', { operation: 'verify', input: { requestId: 'pptx-verify', operation: 'verify', artifactRef: formattedRef } });
if (verified.error) throw new Error(`verify failed: ${verified.error}`);
const verification = verified.value.result ?? verified.value;
console.log(`verify     : ok=${verification.ok} partial=${verification.partial} ${JSON.stringify(verification.summary)}`);
for (const check of verification.checks ?? []) console.log(`  [${check.status}] ${check.id}: ${check.message}`);

// 6. 落盘
const pptx = join(OUT_DIR, '嵌入式的发展路径-配套PPT.pptx');
copyFileSync(localPath(formattedRef.uri), pptx);
console.log(`\npptx       : ${pptx}\n  sha256   : ${sha256(pptx)}  ${readFileSync(pptx).length}B`);

// 7. 用插件自带 LibreOffice 渲染成 PDF，再出逐页 PNG（PPTX 没有内置渲染模块）
const { spawnSync } = await import('node:child_process');
const profile = '-env:UserInstallation=file:///C:/Users/AA/AppData/Local/Temp/lo-profile-pptx';
const convert = spawnSync(SOFFICE, [profile, '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', OUT_DIR, pptx], { encoding: 'utf8', windowsHide: true, timeout: 300000 });
console.log(`render     : soffice exit=${convert.status} ${(convert.stdout ?? '').split('\n').filter(Boolean).slice(-1)[0] ?? ''}`);
const pdf = join(OUT_DIR, '嵌入式的发展路径-配套PPT.pdf');
if (!existsSync(pdf)) throw new Error('LibreOffice did not produce a PDF');
console.log(`pdf        : ${pdf}\n  sha256   : ${sha256(pdf)}  ${statSync(pdf).size}B`);
const raster = spawnSync(PDFTOPPM, ['-r', '110', '-png', pdf, join(OUT_DIR, 'slide')], { encoding: 'utf8', windowsHide: true, timeout: 300000 });
console.log(`png        : pdftoppm exit=${raster.status}`);

writeFileSync(join(OUT_DIR, 'pptx-run-summary.json'), JSON.stringify({
  server: initialize.result?.serverInfo,
  draft: { path: DRAFT, sha256: draft.sha256, sizeBytes: draft.sizeBytes },
  inspection, formatting: {
    formattedRuns: formatting.formattedRuns, formattedSlides: formatting.formattedSlides,
    verifiedFormattingRuns: formatting.verifiedFormattingRuns,
    contrastAdjustedSlides: formatting.contrastAdjustedSlides, titleColorPreservedSlides: formatting.titleColorPreservedSlides,
  },
  verification, deliverable: { path: pptx, sha256: sha256(pptx), sizeBytes: readFileSync(pptx).length },
  pdf: { path: pdf, sha256: sha256(pdf), sizeBytes: statSync(pdf).size },
}, null, 2));

child.kill();
process.exit(0);

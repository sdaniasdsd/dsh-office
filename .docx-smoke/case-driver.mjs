// Drives the freshly built DSH docx MCP server over stdio and runs the
// docx-format-preservation-v1 case against it, without needing a DSH restart.
// The point of the run: restyle in place with docx-edit.formatParagraph and show
// that the comments and footnotes parts survive, which a docx-create rebuild
// cannot do.
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
const SOURCE_PATH = process.argv[6] ?? join(WORKSPACE, 'rich.docx');

mkdirSync(DATA, { recursive: true });
const tmp = join(DATA, '_tmp');
mkdirSync(tmp, { recursive: true });

const child = spawn(process.execPath, [SERVER], {
  env: {
    ...process.env,
    THE_LAST_DOCX_WORKSPACE: WORKSPACE,
    THE_LAST_DOCX_DATA: DATA,
    DOCX_PYTHON: join(RUNTIME, 'python', 'python.exe'),
    DOCX_SOFFICE: join(RUNTIME, 'libreoffice', 'program', 'soffice.com'),
    DOCX_PDFTOPPM: join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
    TEMP: tmp,
    TMP: tmp,
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
    if (message.id !== undefined && pending.has(message.id)) {
      pending.get(message.id)(message);
      pending.delete(message.id);
    }
  }
});
const stderr = [];
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => stderr.push(chunk));

function request(method, params) {
  const id = nextId++;
  const promise = new Promise((resolve) => pending.set(id, resolve));
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`);
  return promise;
}
function notify(method, params) {
  child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
}

/** Tool results carry a JSON string in content[0].text; unwrap or report the error. */
function payload(response) {
  const result = response.result;
  if (!result) return { error: response.error ?? 'no result' };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
}

async function callTool(name, args) {
  return payload(await request('tools/call', { name, arguments: args }));
}

const failures = [];
function check(label, ok, detail = '') {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`);
  if (!ok) failures.push(label);
}

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'case-driver', version: '1' } });
notify('notifications/initialized');
const listed = await request('tools/list', {});
console.log(`server tools: ${(listed.result?.tools ?? []).map((tool) => tool.name).join(', ')}`);
console.log(`serverInfo: ${JSON.stringify(initialize.result?.serverInfo)}`);

// 1. Import the source document.
const imported = await callTool('docx_import', { path: SOURCE_PATH });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const source = imported.value;
console.log(`source: ${source.id} ${source.sizeBytes}B`);
check('file integrity: source sha256 matches the fixture', source.sha256 === 'dc684e35ffd61acb5f9ac96fdaf0492f5ce5d4b1ac1a076e0c4a9817e0899ede');

// 2. Parse it to get real dual-IR anchors.
const analyzed = await callTool('docx_analyze', { artifactRef: source, requestId: 'case-fp-v2-analyze' });
if (analyzed.error) throw new Error(`analyze failed: ${analyzed.error}`);
const dualRef = analyzed.value.bridges.find((bridge) => bridge.kind === 'dual-ir').artifactRef;
const dualRead = await callTool('docx_read_artifact', { artifactRef: dualRef, length: 200000 });
const content = JSON.parse(dualRead.value.text).payload.result.ir.content;
const blocks = content.semantic.blocks;
const byText = (text, kind) => blocks.find((block) => block.text === text && (!kind || block.kind === kind));
const chapter = byText('Chapter One', 'heading');
const section = byText('Section', 'heading');
const intro = byText('Intro text', 'paragraph');
const closing = byText('Closing line', 'paragraph');
check('dual IR exposes the fixture blocks', Boolean(chapter && section && intro && closing));
check('source dual IR carries both annotations', content.semantic.annotations.length === 2,
  `kinds=${content.semantic.annotations.map((a) => a.kind).join(',')}`);

const target = (block) => ({ semanticId: block.id, anchor: block.anchor });

// 3a. Define the named styles FIRST. The edit below writes w:pStyle
//     references; a definition has to exist for those to resolve to something
//     this document controls rather than to whatever the renderer synthesizes.
const edge = (color) => ({ style: 'single', size: 0.5, color, space: 0 });
const stylePlan = {
  styles: [
    {
      styleId: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      paragraph: { outlineLevel: 0, keepNext: true, spacing: { before: 0, after: 10 }, bottomBorder: { style: 'single', size: 0.5, color: '#1F3864', space: 3 } },
      run: { bold: true, size: 18, color: '#1F3864', font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'SimHei' } },
    },
    {
      styleId: 'Heading2', name: 'heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      paragraph: { outlineLevel: 1, keepNext: true, spacing: { before: 16, after: 7 } },
      run: { bold: true, size: 13, color: '#2E5496', font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'SimHei' } },
    },
    {
      styleId: 'TableGrid', type: 'table', name: 'Table Grid',
      table: {
        // docx4j's own Table Grid carries `w:pPr/w:spacing` (after 0, line 240
        // auto) to control the rhythm of the paragraphs inside its cells, and
        // cell margins of 108 twips left/right, 0 top/bottom.
        width: 451, indent: 0, cellMargins: { top: 0, left: 5.4, bottom: 0, right: 5.4 },
        borders: { top: edge('#9AA6B2'), left: edge('#9AA6B2'), bottom: edge('#9AA6B2'), right: edge('#9AA6B2'), insideH: edge('#9AA6B2'), insideV: edge('#9AA6B2') },
      },
      paragraph: { spacing: { after: 0, line: 12, lineRule: 'auto' } },
    },
  ],
};
const styled = await callTool('docx_call', { moduleId: 'docx-styles', operation: 'execute', input: { artifactRef: source, requestId: 'case-fp-v3-styles', plan: stylePlan } });
if (styled.error) throw new Error(`styles failed: ${styled.error}`);
console.log(`styles: ${styled.value.result.defined.map((entry) => `${entry.styleId}=${entry.action}(${entry.properties.length})`).join(', ')}`);
console.log(`styles verification: ${styled.value.result.verification.ok ? 'ok' : 'FAILED'} ${styled.value.result.verification.summary.passed}/${styled.value.result.verification.summary.total}`);
const base = styled.value.result.artifact;

// 3b. Restyle in place: hierarchy, rhythm, and a bold header row.
const plan = {
  edits: [
    { kind: 'formatParagraph', target: target(chapter), styleId: 'Heading1', outlineLevel: 1, spaceAfter: 12, font: { bold: true, size: 18, color: '#1F3864' } },
    { kind: 'formatParagraph', target: target(section), styleId: 'Heading2', outlineLevel: 2, spaceBefore: 12, spaceAfter: 6, font: { bold: true, size: 14, color: '#2E5496' } },
    { kind: 'formatParagraph', target: target(intro), spaceAfter: 8, lineSpacing: 15 },
    { kind: 'formatParagraph', target: target(closing), spaceAfter: 8, lineSpacing: 15 },
  ],
};
const table = blocks.find((block) => block.kind === 'table');
for (const cell of [table.rows[0].cells[0], table.rows[0].cells[1]]) {
  plan.edits.push({ kind: 'formatParagraph', target: target(cell.paragraphs[0]), font: { bold: true }, spaceAfter: 4 });
}
// The table's own geometry. A named style says what a table may look like; the
// table element decides the layout, and the fixture's grid carries no preferred
// widths at all - which is why its cells wrap.
plan.edits.push({
  kind: 'formatTable',
  target: { semanticId: table.id, anchor: table.anchor },
  styleId: 'TableGrid',
  width: 451,
  layout: 'fixed',
  columnWidths: [225.5, 225.5],
  cellMargins: { top: 0, left: 5.4, bottom: 0, right: 5.4 },
  borders: { top: edge('#9AA6B2'), left: edge('#9AA6B2'), bottom: edge('#9AA6B2'), right: edge('#9AA6B2'), insideH: edge('#9AA6B2'), insideV: edge('#9AA6B2') },
  headerRow: true,
  cellVerticalAlignment: 'center',
});
const edited = await callTool('docx_call', { moduleId: 'docx-edit', operation: 'execute', input: { artifactRef: base, requestId: 'case-fp-v3-restyle', plan } });
if (edited.error) throw new Error(`edit failed: ${edited.error}`);
const result = edited.value;
console.log(`edit: ${result.result.edits.length} edits, verification ${result.result.verification.ok ? 'ok' : 'FAILED'}`);
for (const edit of result.result.edits) console.log(`  ${edit.kind} changed=${edit.changed} ${edit.details?.applied ?? ''}`);
for (const check of result.result.verification.checks.filter((entry) => entry.id.startsWith('table.'))) {
  console.log(`  ${check.status === 'pass' ? 'PASS' : 'FAIL'}  ${check.id}: ${check.message}`);
}
check('table grid declares preferred widths (no content-width fallback)',
  result.result.verification.checks.find((entry) => entry.id === 'table.1.grid')?.status === 'pass');
check('table geometry survived the save',
  result.result.verification.checks.find((entry) => entry.id === 'table.1.geometry')?.status === 'pass');
const artifact = result.result.artifact;
console.log(`artifact: ${artifact.sha256} ${artifact.sizeBytes}B`);

// 4. Inspect the output package: the annotation parts must still be there.
const inspected = await callTool('docx_call', { moduleId: 'docx-inspect', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-fp-v2-inspect' } });
if (inspected.error) throw new Error(`inspect failed: ${inspected.error}`);
const parts = inspected.value.result.ir.parts.map((part) => part.name);
console.log(`output parts: ${parts.join(', ')}`);
check('comments part survived', parts.includes('word/comments.xml'));
check('footnotes part survived', parts.includes('word/footnotes.xml'));

// 5. Re-parse to confirm content and annotations, then render.
const reparsed = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-fp-v2-reparse' } });
if (reparsed.error) throw new Error(`reparse failed: ${reparsed.error}`);
const outputContent = reparsed.value.result.ir.content;
const outputBlocks = outputContent.semantic.blocks;
const texts = outputBlocks.filter((block) => block.kind !== 'table').map((block) => block.text);
check('all body sentences preserved',
  ['Chapter One', 'Intro text', 'Commented paragraph', 'Footnote paragraph', 'Section', 'Closing line'].every((text) => texts.includes(text)),
  texts.join(' | '));
const outputTable = outputBlocks.find((block) => block.kind === 'table');
check('table cells still A1,B1,A2,B2',
  outputTable.rows.map((row) => row.cells.map((cell) => cell.text).join(',')).join('/') === 'A1,B1/A2,B2');
check('both annotations still anchored', outputContent.semantic.annotations.length === 2,
  `kinds=${outputContent.semantic.annotations.map((a) => `${a.kind}:${a.text}`).join(',')}`);
check('heading hierarchy applied',
  outputBlocks.find((block) => block.text === 'Chapter One')?.styleId === 'Heading1'
  && outputBlocks.find((block) => block.text === 'Section')?.styleId === 'Heading2');

// 6. Audit the finished package: every reference must land on a definition
//    this document owns, and no property may be out of ECMA-376 sequence.
const audit = await callTool('docx_call', { moduleId: 'docx-styles', operation: 'verify', input: { artifactRef: artifact, requestId: 'case-fp-v3-audit', expectations: { styles: stylePlan.styles } } });
if (audit.error) throw new Error(`styles verify failed: ${audit.error}`);
for (const check of audit.value.verification.checks) console.log(`  ${check.status === 'pass' ? 'PASS' : 'FAIL'}  ${check.id}: ${check.message}`);
check('no dangling style reference remains', audit.value.verification.ok);

const rendered = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-fp-v2-render', config: { limits: { dpi: 144 } } } });
if (rendered.error) throw new Error(`render failed: ${rendered.error}`);
check('renders without error', rendered.value.result.pageCount >= 1, `pages=${rendered.value.result.pageCount}`);
console.log(`page image: ${rendered.value.result.pages[0].image.uri}`);
console.log(`pdf: ${rendered.value.result.pdf.uri}`);

console.log(`\nsource artifact kept for comparison: ${source.uri}`);
if (stderr.length) console.log(`\nserver stderr (tail):\n${stderr.join('').slice(-2000)}`);
child.kill();
console.log(failures.length === 0 ? '\nALL CHECKS PASSED' : `\nFAILED: ${failures.join('; ')}`);
process.exit(failures.length === 0 ? 0 : 1);

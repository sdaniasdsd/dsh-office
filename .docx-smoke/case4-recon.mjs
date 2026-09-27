// Recon for case docx-labor-contract-amendment-v1: import the fixture, parse it
// through the real MCP surface, and dump a digest to disk so the edit plan can
// be built against real anchors rather than guessed strings.
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
const SOURCE_PATH = process.argv[6];
const OUT_DIR = process.argv[7];

mkdirSync(DATA, { recursive: true });
mkdirSync(OUT_DIR, { recursive: true });
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
    TEMP: tmp, TMP: tmp,
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
child.stderr.on('data', (chunk) => stderr.push(chunk));

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

/**
 * Read a tool result that was too large to inline. The server spills anything
 * over 200 000 bytes to an artifact, and `docx_read_artifact` truncates at that
 * same cap and reports `nextOffset` in BYTES with a `complete` flag — advancing
 * by character count instead silently drops the tail of every CJK document.
 */
async function readSpilled(ref) {
  let text = '';
  let offset = 0;
  for (;;) {
    const chunk = await callTool('docx_read_artifact', { artifactRef: ref, offset, length: 200000 });
    if (chunk.error) throw new Error(`spill read failed: ${chunk.error}`);
    const payload = chunk.value;
    text += payload.text ?? '';
    if (payload.complete) break;
    if (!(payload.nextOffset > offset)) throw new Error(`spill read made no progress at ${offset}/${payload.totalBytes}`);
    offset = payload.nextOffset;
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    // Never let the raw single-line payload reach the console: Node prints the
    // whole offending line, which is the entire document IR.
    throw new Error(`spilled result is not valid JSON (${text.length} chars): ${error.message.slice(0, 120)}`);
  }
}

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'case4-recon', version: '1' } });
notify('notifications/initialized');

const imported = await callTool('docx_import', { path: SOURCE_PATH });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const blank = imported.value;
console.log(`source: ${blank.sha256} ${blank.sizeBytes}B`);

const parsed = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: blank, requestId: 'case-04-recon' } });
if (parsed.error) throw new Error(`parse failed: ${parsed.error}`);
const parsedValue = parsed.value.resultRef ? await readSpilled(parsed.value.resultRef) : parsed.value;
const content = parsedValue.result.ir.content;

const blocks = content.semantic.blocks;
const annotations = content.semantic.annotations ?? [];
console.log(`blocks=${blocks.length} annotations=${annotations.length}`);
console.log(`annotations: ${annotations.map((a) => `${a.kind}:${(a.text ?? '').slice(0, 40)}`).join(' | ')}`);

const outline = [];
const walk = (list, prefix) => {
  list.forEach((block, index) => {
    const at = `${prefix}${index}`;
    if (block.kind === 'table') {
      outline.push(`${at}  [TABLE ${block.rows.length}x${block.rows[0]?.cells.length}] path=${block.anchor.structuralPath}`);
      block.rows.forEach((row, ri) => {
        row.cells.forEach((cell, ci) => {
          cell.paragraphs.forEach((paragraph, pi) => {
            outline.push(`${at}.r${ri}c${ci}p${pi}  ${paragraph.kind}  ${(paragraph.text ?? '').slice(0, 120)}`);
          });
        });
      });
    } else {
      outline.push(`${at}  ${block.kind}${block.styleId ? ` style=${block.styleId}` : ''}  ${(block.text ?? '').slice(0, 140)}`);
    }
  });
};
walk(blocks, '');

writeFileSync(join(OUT_DIR, 'recon-blocks.json'), JSON.stringify({ blocks, annotations, sourceMap: content.sourceMap }, null, 2));
writeFileSync(join(OUT_DIR, 'recon-outline.txt'), `${outline.join('\n')}\n`);
console.log(`wrote recon-blocks.json and recon-outline.txt (${outline.length} lines)`);
console.log(`\nfirst lines:\n${outline.slice(0, 12).join('\n')}`);

if (stderr.length) console.log(`\nserver stderr tail:\n${stderr.join('').slice(-1500)}`);
child.kill();

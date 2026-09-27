// Compare the *declared* run sizes of two documents, through the parse module's
// own observation channel: no rendering involved, so a difference here is a fact
// about the files, not about a renderer.
import { spawn } from 'node:child_process';
import { readFileSync, mkdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const DOCS = [];
for (let i = 5; i + 1 < process.argv.length; i += 2) DOCS.push({ path: process.argv[i], label: process.argv[i + 1] });

const tmp = join(DATA, '_tmp');
mkdirSync(tmp, { recursive: true });
mkdirSync(WORKSPACE, { recursive: true });
const child = spawn(process.execPath, [SERVER], {
  env: { ...process.env, THE_LAST_DOCX_WORKSPACE: WORKSPACE, THE_LAST_DOCX_DATA: DATA, TEMP: tmp, TMP: tmp },
  stdio: ['pipe', 'pipe', 'pipe'],
});
child.stderr.setEncoding('utf8');
child.stderr.on('data', (chunk) => process.stderr.write(chunk));

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
const request = (method, params) => { const id = nextId++; const promise = new Promise((resolve) => pending.set(id, resolve)); child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`); return promise; };
const notify = (method, params) => child.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', method, params })}\n`);
function payload(response) {
  const result = response.result;
  if (!result) return { error: response.error ?? 'no result' };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
}
const callTool = async (name, args) => payload(await request('tools/call', { name, arguments: args }));

/** A module result too large for the tool reply is spilled to an artifact. */
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

/** Samples and in-table counts are collected beside the counts, per document. */
let SAMPLE_BY_SIZE = new Map();
let TABLE_BY_SIZE = new Map();
const samplesOf = (_counts, size) => {
  if (!SAMPLE_BY_SIZE.has(size)) SAMPLE_BY_SIZE.set(size, []);
  return SAMPLE_BY_SIZE.get(size);
};
const inTable = (paragraph) => /\/w:tbl\[/u.test(String(paragraph.pointer ?? ''));

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'artifact-sizing', version: '1' } });
notify('notifications/initialized');

for (const doc of DOCS) {
  const imported = await callTool('docx_import', { path: doc.path });
  if (imported.error) throw new Error(`import ${doc.path} failed: ${imported.error}`);
  const ref = imported.value;
  const parsed = await callTool('docx_call', {
    moduleId: 'docx-parse', operation: 'execute',
    input: { artifactRef: ref, requestId: `sizing-${doc.label}`, baseRevision: ref.revision, options: { featureFlags: { parseFormatting: true } } },
  });
  if (parsed.error) throw new Error(`parse ${doc.label} failed: ${parsed.error}`);
  const parsedResult = parsed.value?.result?.ir ? parsed.value.result : (parsed.value?.resultRef ? await readSpilled(parsed.value.resultRef) : parsed.value?.result);
  const formatting = (parsedResult?.ir ?? parsedResult?.result?.ir)?.content?.formatting;
  if (!formatting) {
    console.log(`\n[debug] parse ${doc.label} payload keys: ${JSON.stringify(Object.keys(parsedResult ?? {}))}`);
    throw new Error(`no formatting observation for ${doc.label}`);
  }

  const bySize = new Map();
  const byFace = new Map();
  const noSize = [];
  SAMPLE_BY_SIZE = new Map();
  TABLE_BY_SIZE = new Map();
  for (const paragraph of formatting.paragraphs) {
    let anySize = false;
    const text = paragraph.runs.map((run) => run.text).join('').trim().slice(0, 22);
    for (const run of paragraph.runs) {
      if (!run.text.trim()) continue;
      if (run.size !== undefined) {
        anySize = true;
        bySize.set(run.size, (bySize.get(run.size) ?? 0) + 1);
        const samples = samplesOf(bySize, run.size);
        if (samples.length < 4 && text && !samples.includes(text)) samples.push(text);
        if (inTable(paragraph)) TABLE_BY_SIZE.set(run.size, (TABLE_BY_SIZE.get(run.size) ?? 0) + 1);
      }
      if (run.eastAsia) byFace.set(run.eastAsia, (byFace.get(run.eastAsia) ?? 0) + 1);
    }
    if (!anySize && text) noSize.push(text);
  }
  const size = readFileSync(doc.path).length;
  console.log(`\n===== ${doc.label} =====`);
  console.log(`path        : ${doc.path}`);
  console.log(`sha256      : ${createHash('sha256').update(readFileSync(doc.path)).digest('hex')}  ${size}B`);
  console.log(`paragraphs  : ${formatting.paragraphs.length} observed, ${noSize.length} declare no size`);
  if (process.env.SHOW_POINTERS === '1') {
    for (const paragraph of formatting.paragraphs.slice(0, 3)) console.log(`  entry: ${JSON.stringify(paragraph).slice(0, 260)}`);
    const cellEntry = formatting.paragraphs.find((entry) => /\/w:tbl\[/u.test(String(entry.pointer ?? '')));
    if (cellEntry) console.log(`  cell entry: ${JSON.stringify(cellEntry).slice(0, 260)}`);
    if (process.env.SHOW_CELLS === '1') {
      const cells = formatting.paragraphs.filter((entry) => /\/w:tbl\[/u.test(String(entry.pointer ?? '')));
      console.log(`  --- ${cells.length} in-table paragraphs ---`);
      for (const entry of cells) {
        const run = entry.runs.find((candidate) => candidate.text.trim()) ?? entry.runs[0] ?? {};
        console.log(`      ${String(entry.pointer).replace('word/document.xml!/w:document/w:body/', '')}  size=${run.size ?? '-'} face=${run.eastAsia ?? '-'} bold=${run.bold ?? '-'}  ${String(run.text ?? '').trim().slice(0, 18)}`);
      }
    }
  }
  for (const [value, count] of [...bySize].sort((a, b) => b[0] - a[0])) {
    const inTables = TABLE_BY_SIZE.get(value) ?? 0;
    console.log(`  declared size ${String(value).padStart(6)} pt  ×${String(count).padStart(4)}  (${inTables} inside tables)   e.g. ${(SAMPLE_BY_SIZE.get(value) ?? []).join(' | ')}`);
  }
  for (const [face, count] of [...byFace].sort((a, b) => b[1] - a[1])) console.log(`  declared face ${face.padEnd(8)} ×${String(count).padStart(4)}`);
  if (noSize.length) console.log(`  declare-no-size ×${noSize.length} e.g. ${noSize.slice(0, 4).join(' | ')}`);
}

child.kill();
process.exit(0);

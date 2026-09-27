// Re-run the labour contract through the new reference path: take the approved
// standard document as the reference, translate it into a creation plan, and
// build the result — then render it so the two page layouts can be compared.
import { spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
const REFERENCE_PATH = process.argv[6];
const OUT_DIR = process.argv[7];

mkdirSync(DATA, { recursive: true });
mkdirSync(WORKSPACE, { recursive: true });
mkdirSync(OUT_DIR, { recursive: true });
const tmp = join(DATA, '_tmp');
mkdirSync(tmp, { recursive: true });
const sha256File = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');
const localPath = (uri) => decodeURIComponent(String(uri).replace(/^file:\/\/\//, '').replace(/^file:\/\//, '').replace(/^file:/, ''));

const child = spawn(process.execPath, [SERVER], {
  env: {
    ...process.env,
    THE_LAST_DOCX_WORKSPACE: WORKSPACE,
    THE_LAST_DOCX_DATA: DATA,
    ...(existsSync(join(RUNTIME, 'python', 'python.exe')) ? { DOCX_PYTHON: join(RUNTIME, 'python', 'python.exe') } : {}),
    ...(existsSync(join(RUNTIME, 'libreoffice', 'program', 'soffice.com')) ? { DOCX_SOFFICE: join(RUNTIME, 'libreoffice', 'program', 'soffice.com') } : {}),
    ...(existsSync(join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe')) ? { DOCX_PDFTOPPM: join(RUNTIME, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe') } : {}),
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
child.stderr.on('data', (chunk) => { stderr.push(chunk); process.stderr.write(chunk); });
child.on('exit', (code, signal) => { if (code !== 0) process.stderr.write(`\n[server exited: code=${code} signal=${signal}]\n`); });
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

const failures = [];
const check = (label, ok, detail = '') => { console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`); if (!ok) failures.push(label); };

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'ref-gen', version: '1' } });
notify('notifications/initialized');
const listed = await request('tools/list', {});
console.log(`serverInfo: ${JSON.stringify(initialize.result?.serverInfo)}`);
console.log(`tools: ${(listed.result?.tools ?? []).map((t) => t.name).join(', ')}\n`);
check('the reference path is exposed as a tool', (listed.result?.tools ?? []).some((t) => t.name === 'docx_from_reference'));

const referenceSha = sha256File(REFERENCE_PATH);
const imported = await callTool('docx_import', { path: REFERENCE_PATH });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const reference = imported.value;
console.log(`reference: ${reference.sha256.slice(0, 12)}… ${reference.sizeBytes}B`);

const built = await callTool('docx_from_reference', {
  referenceRef: reference,
  requestId: 'contract-from-standard',
  options: {
    preset: 'chinese-contract',
    scenario: 'formal-record',
    page: { size: 'A4', marginsMm: { top: 25, right: 28, bottom: 25, left: 30 } },
    pageNumberStyle: 'pageOfTotal',
    tableWidthMm: 152,
    // Ask the parser for the presentational facts, so an entirely bold paragraph
    // becomes a heading and a centred one keeps its centring.
    carryFormatting: true,
  },
});
if (built.error) throw new Error(`docx_from_reference failed: ${built.error}`);

const report = built.value.report;
const created = built.value.created;
const artifact = created.result.artifactRef;
const design = created.result.design;
console.log(`\ncarried      : ${JSON.stringify(report.carried)}`);
console.log(`skipped      : ${report.skipped.length ? JSON.stringify(report.skipped) : '(none)'}`);
console.log(`not carried  : ${JSON.stringify(report.notCarriedByDesign)}`);
console.log(`unmapped     : ${Object.keys(report.unmappedStyles).length ? JSON.stringify(report.unmappedStyles) : '(none)'}`);
for (const note of report.notes) console.log(`note         : ${note}`);
console.log(`\ndesign       : register=${design?.register} decoration=${design?.decoration} source=${design?.source}`);
console.log(`artifact     : ${artifact.sha256} ${artifact.sizeBytes}B`);

check('intent_complete: the reference produced a new artifact', Boolean(artifact?.sha256) && artifact.uri !== reference.uri);
check('the reference was never modified', sha256File(REFERENCE_PATH) === referenceSha);
check('no block of the reference was silently dropped', report.skipped.length === 0, JSON.stringify(report.skipped));
check('nothing a plan cannot carry was present in the reference',
  Object.values(report.notCarriedByDesign).every((count) => count === 0), JSON.stringify(report.notCarriedByDesign));
check('the plan carries the reference’s blocks', report.carried.paragraphs >= 70 && report.carried.tables === 3, JSON.stringify(report.carried));
check('the reference’s hierarchy is carried as far as the observation exposes it',
  report.carried.boldAsHeading > 0,
  `the reference marks its clause headings with direct run formatting and no paragraph style; ${report.carried.boldAsHeading} entirely-bold paragraph(s) became headings`);

const deliverable = join(OUT_DIR, 'labour-contract-from-standard.docx');
copyFileSync(localPath(artifact.uri), deliverable);
console.log(`\ndeliverable: ${deliverable} ${sha256File(deliverable)} ${readFileSync(deliverable).length}B`);

if (process.env.SKIP_MODULE_RENDER === '1') {
  console.log('module render: skipped (SKIP_MODULE_RENDER=1) — render the .docx with the renderer the comparison uses');
} else {
  const rendered = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: artifact, requestId: 'cfs-render', config: { limits: { dpi: 144 } } } });
  if (rendered.error) {
    console.log(`module render: ${rendered.error}`);
    check('complete_render: the module rendered the result', false, rendered.error);
  } else {
    const render = rendered.value.result;
    console.log(`module render: ${render.pageCount} pages`);
    copyFileSync(localPath(render.pdf.uri), join(OUT_DIR, 'labour-contract-from-standard.pdf'));
    for (const page of render.pages ?? []) {
      copyFileSync(localPath(page.image.uri), join(OUT_DIR, `from-standard-page-${page.index ?? 1}.png`));
    }
    check('complete_render: the module rendered the result', render.pageCount >= 1, `pages=${render.pageCount}`);
  }
}

writeFileSync(join(OUT_DIR, 'run-summary.json'), JSON.stringify({
  reference: { path: REFERENCE_PATH, sha256: referenceSha, sizeBytes: readFileSync(REFERENCE_PATH).length },
  deliverable: { path: deliverable, sha256: sha256File(deliverable), sizeBytes: readFileSync(deliverable).length },
  report, design, failures,
}, null, 2));

if (stderr.length) console.log(`\nserver stderr tail:\n${stderr.join('').slice(-800)}`);
child.kill();
console.log(failures.length === 0 ? '\nALL CHECKS PASSED' : `\nFAILED: ${failures.join('; ')}`);
process.exit(failures.length === 0 ? 0 : 1);

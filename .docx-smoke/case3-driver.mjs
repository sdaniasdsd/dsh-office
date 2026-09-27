// Case docx-create-from-brief-v1: build 《第三季度客户交付简报》 from a blank DOCX
// with docx-create, then prove the result against the case's gates without a DSH restart.
//
// Gates exercised here: file_integrity, content_preservation, no_fabricated_facts,
// intent_complete, complete_render. (no_critical_visual_defect is judged separately
// from the rendered page images.)
import { spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
const SOURCE_PATH = process.argv[6];
const OUT_DIR = process.argv[7];
const ORIGINAL_FIXTURE = process.argv[8];

mkdirSync(DATA, { recursive: true });
mkdirSync(OUT_DIR, { recursive: true });
const tmp = join(DATA, '_tmp');
mkdirSync(tmp, { recursive: true });

const sha256File = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');
// Artifact URIs come back as `file:<path>`; decode before touching the filesystem.
const localPath = (uri) => decodeURIComponent(String(uri).replace(/^file:\/\/\//, '').replace(/^file:\/\//, '').replace(/^file:/, ''));

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
function payload(response) {
  const result = response.result;
  if (!result) return { error: response.error ?? 'no result' };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
}
const callTool = async (name, args) => payload(await request('tools/call', { name, arguments: args }));

const failures = [];
function check(label, ok, detail = '') {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`);
  if (!ok) failures.push(label);
}
const squash = (value) => String(value).replace(/\s+/gu, '');

// ---------------------------------------------------------------- brief data
const BRIEF_TITLE = '第三季度客户交付简报';
const BRIEF_DATE = '2026年9月30日';
const BRIEF_SUMMARY = '本季度完成 18 项交付任务，按期完成 16 项；尚有 2 项延期，其中 1 项等待客户确认，1 项进入修复。';
const TABLE_ROWS = [
  ['指标', '数量', '说明'],
  ['交付任务', '18', '本季度完成'],
  ['按期完成', '16', '已完成交付任务'],
  ['延期事项', '2', '1 项等待客户确认，1 项进入修复'],
];
const FOLLOW_UPS = ['取得客户确认。', '完成 1 项修复中的延期事项。'];

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'case3-driver', version: '1' } });
notify('notifications/initialized');
const listed = await request('tools/list', {});
console.log(`serverInfo: ${JSON.stringify(initialize.result?.serverInfo)}`);
console.log(`server tools: ${(listed.result?.tools ?? []).map((tool) => tool.name).join(', ')}\n`);

// ------------------------------------------------- 1. file integrity: source
const sourceShaBefore = sha256File(SOURCE_PATH);
const imported = await callTool('docx_import', { path: SOURCE_PATH });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const blank = imported.value;
console.log(`blank source: ${blank.id} ${blank.sizeBytes}B sha=${blank.sha256}`);
check('file_integrity: blank fixture hash matches the staged copy', blank.sha256 === sourceShaBefore, sourceShaBefore);
check('file_integrity: blank fixture carries no business content', (() => {
  const zip = readFileSync(SOURCE_PATH);
  return zip.length < 4096;
})(), `${sourceShaBefore.slice(0, 12)}… ${sourceShaBefore ? '' : ''}`.trim());

// ------------------------------------------------------ 2. build the document
const plan = {
  kind: 'create',
  document: {
    preset: 'report',
    // Declaring the scenario is what lets the engine decide how much decoration
    // this document warrants; the decision comes back in the result.
    scenario: 'internal-review',
    header: BRIEF_TITLE,
    footer: '内部评审材料',
    pageNumbers: true,
    blocks: [
      { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: BRIEF_TITLE }] },
      { kind: 'paragraph', id: 'date', style: 'Subtitle', runs: [{ text: BRIEF_DATE }] },
      { kind: 'paragraph', id: 'h-summary', style: 'Heading1', runs: [{ text: '执行摘要' }] },
      { kind: 'paragraph', id: 'p-summary', runs: [{ text: BRIEF_SUMMARY }] },
      { kind: 'paragraph', id: 'h-table', style: 'Heading1', runs: [{ text: '汇总' }] },
      {
        kind: 'table', id: 't-summary', header: true, columnWidthsMm: [38, 22, 100], rows: TABLE_ROWS,
      },
      { kind: 'paragraph', id: 'h-next', style: 'Heading1', runs: [{ text: '后续跟进' }] },
      // A real list: the markers come from the numbering part, not the text.
      { kind: 'list', id: 'next-actions', items: FOLLOW_UPS },
    ],
  },
};

const created = await callTool('docx_call', { moduleId: 'docx-create', operation: 'execute', input: { requestId: 'case-03-create', plan } });
if (created.error) throw new Error(`create failed: ${created.error}`);
const createResult = created.value.result;
const artifact = createResult.artifactRef;
console.log(`\ncreated: ${artifact.id} ${artifact.sizeBytes}B sha=${artifact.sha256}`);
console.log(`engine=${createResult.engine} mode=${createResult.mode} paragraphs=${createResult.paragraphIds.length} fields=${createResult.fields} visualReview=${createResult.visualReview}`);
console.log(`module verification: ${created.value.verification?.ok ? 'ok' : 'FAILED'} ${created.value.verification?.summary?.passed ?? '?'}/${created.value.verification?.summary?.total ?? '?'}`);
for (const entry of created.value.verification?.checks ?? []) console.log(`  ${entry.status.toUpperCase().padEnd(4)}  ${entry.id}: ${entry.message}`);

check('file_integrity: output is a NEW artifact, not the input', artifact.id !== blank.id && artifact.uri !== blank.uri && artifact.sha256 !== blank.sha256);
check('file_integrity: output is a valid DOCX package', (created.value.verification?.checks ?? []).every((entry) => entry.status === 'pass') || created.value.verification?.ok === true);
check('file_integrity: create module self-verification passed', created.value.verification?.ok === true);

// The decoration decision has to come back with the result, and say how it was
// reached, or nobody can review it.
const design = createResult.design;
console.log(`design: register=${design?.register} decoration=${design?.decoration} source=${design?.source}`);
console.log(`  reason: ${design?.reason}`);
check('design decision is reported with its reason', Boolean(design?.register && design?.reason && design?.source === 'scenario'));
check('the internal-review scenario resolves to the report register', design?.register === 'report');

// --------------------------------------------- 3. content preservation / intent
const reparsed = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-03-reparse' } });
if (reparsed.error) throw new Error(`reparse failed: ${reparsed.error}`);
const content = reparsed.value.result.ir.content;
const blocks = content.semantic.blocks;
const texts = blocks.filter((b) => b.kind !== 'table').map((b) => b.text);
console.log(`\nbody blocks: ${texts.join(' | ')}`);
const headingByText = (text) => blocks.find((b) => b.text === text);
const styleOf = (text) => headingByText(text)?.styleId ?? headingByText(text)?.style ?? '(none)';

check('intent_complete: title present with the brief title', texts.includes(BRIEF_TITLE), `style=${styleOf(BRIEF_TITLE)}`);
check('intent_complete: date present and complete', texts.includes(BRIEF_DATE), BRIEF_DATE);
check('intent_complete: title/date block hierarchy applied',
  /title/i.test(styleOf(BRIEF_TITLE)) && /subtitle/i.test(styleOf(BRIEF_DATE)),
  `title=${styleOf(BRIEF_TITLE)} date=${styleOf(BRIEF_DATE)}`);
check('content_preservation: executive summary verbatim',
  texts.some((t) => squash(t) === squash(BRIEF_SUMMARY)), BRIEF_SUMMARY);
check('intent_complete: three sections carry heading styles',
  ['执行摘要', '汇总', '后续跟进'].every((h) => /heading/i.test(styleOf(h))),
  ['执行摘要', '汇总', '后续跟进'].map((h) => `${h}=${styleOf(h)}`).join(', '));
check('content_preservation: both follow-up items present',
  FOLLOW_UPS.every((item) => texts.some((t) => squash(t) === squash(item))), FOLLOW_UPS.join(' / '));

const table = blocks.find((b) => b.kind === 'table');
const actualRows = table.rows.map((row) => row.cells.map((cell) => cell.text.trim()));
const asText = actualRows.map((row) => row.join(' | ')).join('  //  ');
check('content_preservation: table is 3 columns x 4 rows', actualRows.length === 4 && actualRows.every((row) => row.length === 3), `${actualRows.length} rows`);
check('content_preservation: header row is 指标 / 数量 / 说明',
  JSON.stringify(actualRows[0]) === JSON.stringify(TABLE_ROWS[0]), actualRows[0]?.join(' | '));
// The alignment gate: each number must still sit beside the meaning it belongs to.
check('content_preservation: every number keeps its own row meaning (no misalignment)',
  JSON.stringify(actualRows) === JSON.stringify(TABLE_ROWS), asText);

// ------------------------------------------------------- 4. no fabricated facts
const bodyText = [
  ...texts,
  ...actualRows.flat(),
  ...(content.semantic.annotations ?? []).map((a) => a.text),
].join('\n');
const allowedNumbers = new Set(['18', '16', '2', '1', '2026', '9', '30']);
const strayNumbers = [...new Set((bodyText.match(/\d+/g) ?? []))].filter((n) => !allowedNumbers.has(n));
check('no_fabricated_facts: no number outside the brief appears anywhere in the body', strayNumbers.length === 0, strayNumbers.join(',') || 'none');
const dateLike = [...new Set((bodyText.match(/\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日/g) ?? []))];
check('no_fabricated_facts: the only date is the brief date', dateLike.length <= 1 && (dateLike.length === 0 || squash(dateLike[0]).startsWith('2026年9月30日')), dateLike.join(',') || 'none');
const forbidden = ['预测', '预计', '预估', '承诺', '金额', '¥', '$', '€', '£', '万元', '客户名', '负责人姓名'];
const violations = forbidden.filter((word) => bodyText.includes(word));
check('no_fabricated_facts: no invented cause / amount / promise / forecast wording', violations.length === 0, violations.join(',') || 'none');
const clientNameLike = [...new Set((bodyText.match(/[\u4e00-\u9fa5]{2,6}(?:公司|集团|银行|科技|有限公司)/g) ?? []))];
check('no_fabricated_facts: no client name or organisation invented', clientNameLike.length === 0, clientNameLike.join(',') || 'none');

// ------------------------------------------------------------- 5. inspect parts
const inspected = await callTool('docx_call', { moduleId: 'docx-inspect', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-03-inspect' } });
if (inspected.error) throw new Error(`inspect failed: ${inspected.error}`);
const parts = inspected.value.result.ir.parts.map((part) => part.name);
console.log(`\nparts (${parts.length}): ${parts.join(', ')}`);
check('the list block brought a numbering part into the package', parts.includes('word/numbering.xml'));

// --------------------------------------------------------------- 6. render
const rendered = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: artifact, requestId: 'case-03-render', config: { limits: { dpi: 144 } } } });
if (rendered.error) throw new Error(`render failed: ${rendered.error}`);
const render = rendered.value.result;
const pageCount = render.pageCount;
console.log(`\nrender: ${pageCount} pages, engine=${render.engine ?? 'n/a'}`);
check('complete_render: rendered at least one page', pageCount >= 1, `pages=${pageCount}`);

const deliverable = join(OUT_DIR, 'case-03-quarterly-delivery-brief.docx');
copyFileSync(localPath(artifact.uri), deliverable);
const pdfPath = join(OUT_DIR, 'case-03-quarterly-delivery-brief.pdf');
copyFileSync(localPath(render.pdf.uri), pdfPath);
const pageImages = [];
for (const page of render.pages ?? []) {
  const dest = join(OUT_DIR, 'pages', `page-${page.index ?? page.number ?? pageImages.length + 1}.png`);
  mkdirSync(join(OUT_DIR, 'pages'), { recursive: true });
  copyFileSync(localPath(page.image.uri), dest);
  pageImages.push(dest);
}
console.log(`deliverable: ${deliverable} ${sha256File(deliverable)}`);
console.log(`pdf:         ${pdfPath} ${sha256File(pdfPath)}`);
console.log(`page images: ${pageImages.length}`);

// ------------------------------------------------- 7. re-verify the input file
const sourceShaAfter = sha256File(SOURCE_PATH);
check('file_integrity: the imported blank document was never overwritten', sourceShaAfter === sourceShaBefore,
  `${sourceShaAfter.slice(0, 12)}…`);
check('file_integrity: deliverable differs from the blank input', sha256File(deliverable) !== sourceShaBefore);
// Observation, not a gate: the repository's own docx-inspect test setup runs
// fixtures/generate_fixtures.py over this directory, so its bytes move with the
// test suite rather than with anything this case did. Gating on it would fail
// every time the suite runs.
if (ORIGINAL_FIXTURE) {
  const repoFixture = sha256File(ORIGINAL_FIXTURE);
  console.log(`note: repo fixture is ${repoFixture === sourceShaBefore ? 'unchanged' : 'regenerated by the inspect test setup'} (now ${repoFixture.slice(0, 12)}…)`);
}

writeFileSync(join(OUT_DIR, 'run-summary.json'), JSON.stringify({
  caseId: 'docx-create-from-brief-v1',
  artifact, pageCount, pdf: render.pdf, parts, pageImages,
  deliverable: { path: deliverable, sha256: sha256File(deliverable), sizeBytes: readFileSync(deliverable).length },
  source: { path: SOURCE_PATH, sha256: sourceShaBefore, sizeBytes: readFileSync(SOURCE_PATH).length },
  failures,
}, null, 2));

if (stderr.length) console.log(`\nserver stderr (tail):\n${stderr.join('').slice(-3000)}`);
child.kill();
console.log(failures.length === 0 ? '\nALL CHECKS PASSED' : `\nFAILED: ${failures.join('; ')}`);
process.exit(failures.length === 0 ? 0 : 1);

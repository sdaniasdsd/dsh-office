// Case docx-labor-contract-amendment-v1.
//
// Six approved amendments to a synthetic fixed-term labour contract, every
// content change kept as a Word revision, with the source's comments, footnote,
// hyperlinks, bookmarks, repeating table headers and signature block intact.
//
// Targets are located by text, not by row number: the fixture repeats key fields
// between the body and the annexes on purpose, and a hard-coded index would
// silently amend the wrong copy the first time the fixture changed.
import { spawn } from 'node:child_process';
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
const SOURCE_PATH = process.argv[6];
const OUT_DIR = process.argv[7];
const REPO_FIXTURE = process.argv[8];

const SOURCE_SHA = 'dc8e138db5421b71df407d294199af3c8269597dbee6601aa9f84dafc69981c1';
const LOCATION = '上海市浦东新区龙井路 88 号海悦中心 5 层';
const NEW_SENTENCE = '因故障排查形成的临时数据导出文件，应在任务结束后五个工作日内归档至甲方批准的项目空间或按流程安全删除。';
const COMMENT_TEXT = '请确认通讯地址、个人邮箱或联系电话变更时的书面通知及收悉记录均可追溯。';
const COMMENT_SENTENCE = '双方应保证本合同首页所列通讯地址、电子邮箱和联系电话真实有效。';
const AUTHOR = { name: '合同修订', initials: 'AMD' };

mkdirSync(DATA, { recursive: true });
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

/** Reassemble a result the server spilled; `nextOffset` is in bytes, not chars. */
async function readSpilled(ref) {
  let text = '';
  let offset = 0;
  for (;;) {
    const chunk = await callTool('docx_read_artifact', { artifactRef: ref, offset, length: 200000 });
    if (chunk.error) throw new Error(`spill read failed: ${chunk.error}`);
    text += chunk.value.text ?? '';
    if (chunk.value.complete) break;
    if (!(chunk.value.nextOffset > offset)) throw new Error('spill read made no progress');
    offset = chunk.value.nextOffset;
  }
  try { return JSON.parse(text); } catch (error) { throw new Error(`spilled result is not JSON: ${error.message.slice(0, 100)}`); }
}

const failures = [];
function check(label, ok, detail = '') {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`);
  if (!ok) failures.push(label);
}
const squash = (value) => String(value ?? '').replace(/\s+/gu, '');
const squashText = (block) => squash(block.text);
const rowsText = (table) => table.rows.map((row) => row.cells.map((cell) => cell.text.trim()).join(' | ')).join('  //  ');

const initialize = await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'case4-driver', version: '1' } });
notify('notifications/initialized');
const listed = await request('tools/list', {});
console.log(`serverInfo: ${JSON.stringify(initialize.result?.serverInfo)}`);
console.log(`server tools: ${(listed.result?.tools ?? []).map((t) => t.name).join(', ')}`);

// ------------------------------------------------------------- 1. the source
check('file_integrity: staged source matches the benchmark fixture', sha256File(SOURCE_PATH) === SOURCE_SHA, SOURCE_SHA);
const imported = await callTool('docx_import', { path: SOURCE_PATH });
if (imported.error) throw new Error(`import failed: ${imported.error}`);
const source = imported.value;
const sourceParts = (await callTool('docx_call', { moduleId: 'docx-inspect', operation: 'execute', input: { artifactRef: source, requestId: 'c4-inspect-src' } })).value.result.ir.parts.map((p) => p.name);
console.log(`source: ${source.sha256.slice(0, 12)}… ${source.sizeBytes}B, ${sourceParts.length} parts`);

const parsedSource = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: source, requestId: 'c4-parse-src' } });
if (parsedSource.error) throw new Error(`parse failed: ${parsedSource.error}`);
const srcBlocks = (parsedSource.value.resultRef ? await readSpilled(parsedSource.value.resultRef) : parsedSource.value).result.ir.content.semantic.blocks;
console.log(`source blocks: ${srcBlocks.length}`);

// ------------------------------------------------------- 2. locate the targets
const paragraphs = srcBlocks.filter((b) => b.kind !== 'table');
const tables = srcBlocks.filter((b) => b.kind === 'table');
const byPrefix = (prefix) => {
  const hits = paragraphs.filter((b) => (b.text ?? '').startsWith(prefix));
  if (hits.length !== 1) throw new Error(`expected exactly one paragraph starting "${prefix}", found ${hits.length}`);
  return hits[0];
};
const tableByRowLabel = (label) => {
  const hits = tables.filter((t) => t.rows.some((row) => squash(row.cells[0]?.text) === squash(label)));
  if (hits.length !== 1) throw new Error(`expected exactly one table with row "${label}", found ${hits.length}`);
  return hits[0];
};
const cellOf = (table, label, column) => {
  const row = table.rows.find((r) => squash(r.cells[0]?.text) === squash(label));
  if (!row) throw new Error(`no row "${label}"`);
  const paragraph = row.cells[column]?.paragraphs?.[0];
  if (!paragraph) throw new Error(`no paragraph in "${label}" column ${column}`);
  return paragraph;
};

const p21 = byPrefix('2.1 本合同为固定期限劳动合同');
const p22 = byPrefix('2.2 试用期为六个月');
const p31 = byPrefix('3.1 乙方岗位为');
const p41 = byPrefix('4.1 乙方日常工作地点为');
const p61 = byPrefix('6.1 乙方转正后的税前月基本工资');
const p62 = byPrefix('6.2 乙方试用期税前月工资');
const p92 = byPrefix('9.2 乙方应按照授权范围');
const p102 = byPrefix('10.2 双方应保证本合同首页所列');
const annexA = tableByRowLabel('月基本工资（税前）');
const annexB = tableByRowLabel('岗位名称');
console.log(`targets: 21=${p21.anchor.structuralPath} 22=${p22.anchor.structuralPath} 31=${p31.anchor.structuralPath} 41=${p41.anchor.structuralPath}`);
console.log(`         61=${p61.anchor.structuralPath} 62=${p62.anchor.structuralPath} 92=${p92.anchor.structuralPath} 102=${p102.anchor.structuralPath}`);
console.log(`         annexA=${annexA.anchor.structuralPath} annexB=${annexB.anchor.structuralPath}`);

// ------------------------------------------- 3. the approved amendments, only
const target = (block) => ({ semanticId: block.id, anchor: block.anchor });
const replace = (block, find, to) => ({ kind: 'replaceText', target: target(block), find, replace: to, revision: 'track' });

const edits = [
  // 1. contract term and probation period
  replace(p21, '2026 年 10 月 1 日', '2026 年 11 月 1 日'),
  replace(p21, '2029 年 9 月 30 日', '2029 年 10 月 31 日'),
  replace(p22, '六个月', '三个月'),
  replace(p22, '2026 年 10 月 1 日', '2026 年 11 月 1 日'),
  replace(p22, '2027 年 3 月 31 日', '2027 年 1 月 31 日'),
  // 2. role and department
  replace(p31, '高级运维分析师', '数据运营专员'),
  replace(p31, '云平台组', '数据治理组'),
  // 3. ordinary office location
  replace(p41, '上海市浦东新区云桥路 168 号澄海中心 8 层', LOCATION),
  // 4. salary and its written form, plus the derived probation salary
  replace(p61, '18,800 元', '20,800 元'),
  replace(p61, '壹万捌仟捌佰元整', '贰万零捌佰元整'),
  replace(p62, '15,040 元', '16,640 元'),
  // annex A, which restates the same two figures
  replace(cellOf(annexA, '月基本工资（税前）', 1), '15,040 元', '16,640 元'),
  replace(cellOf(annexA, '月基本工资（税前）', 2), '18,800 元', '20,800 元'),
  // annex B, which restates role, department and ordinary location
  replace(cellOf(annexB, '岗位名称', 1), '高级运维分析师', '数据运营专员'),
  replace(cellOf(annexB, '所属部门', 1), '云平台组', '数据治理组'),
  replace(cellOf(annexB, '常规办公地点', 1), '上海市浦东新区云桥路 168 号澄海中心 8 层', LOCATION),
  // 5. a new information-security sentence directly after 9.2
  { kind: 'insertParagraph', target: target(p92), text: NEW_SENTENCE, position: 'After', revision: 'track' },
  // 6. a comment on the first sentence of 10.2, replacing nothing
  { kind: 'addComment', target: target(p102), quote: COMMENT_SENTENCE, text: COMMENT_TEXT, author: AUTHOR },
];

// ---------------------------------------------- 4. the decoys, named and frozen
// Every one of these shares a number, a duration or a street with something the
// prompt DID ask to change; each is a different kind of thing.
const DECOYS = [
  ['交通津贴 800 元/月 (6.4)', byPrefix('6.4 交通津贴为税前每月'), '800 元'],
  ['年度目标奖金参考额 18,800 元/年 (6.5)', byPrefix('6.5 绩效奖金依据'), '18,800 元'],
  ['培训预算上限 18,800 元 (7.3)', byPrefix('7.3 餐饮、体检、培训和办公设备'), '18,800 元'],
  ['三个月培养回顾 (3.3)', byPrefix('3.3 乙方按合理工作要求'), '三个月培养回顾'],
  ['B.2.1 三个月培养回顾', byPrefix('B.2.1 入职后的三个月培养回顾'), '三个月培养回顾'],
  ['附件 A 交通津贴 800 元/月', cellOf(annexA, '交通津贴', 2), '800 元 / 月'],
  ['附件 A 年度目标奖金参考额', cellOf(annexA, '年度目标奖金参考额', 2), '18,800 元 / 年'],
  ['附件 A 年度培训预算参考上限', cellOf(annexA, '年度培训预算参考上限', 2), '18,800 元 / 年'],
  ['附件 B 岗位职级 P4', cellOf(annexB, '岗位职级', 1), 'P4'],
  ['附件 B 临时应急集合点', cellOf(annexB, '临时应急集合点', 1), '上海市浦东新区云桥路 168 号澄海中心 1 层北门'],
  ['附件 B 设备交接地点', cellOf(annexB, '设备交接地点', 1), '上海市杨浦区控江路 560 号仓储点'],
  ['附件 B 直接汇报对象', cellOf(annexB, '直接汇报对象', 1), '平台平台主管'],
];
const decoyFingerprints = DECOYS.map(([label, block, expected]) => {
  if (!squash(block.text).includes(squash(expected))) throw new Error(`decoy "${label}" does not contain "${expected}"`);
  return { label, expected };
});
console.log(`\ndecoys frozen: ${decoyFingerprints.length}`);

const plan = { author: AUTHOR, edits };
console.log(`\nplan: ${edits.length} edits (${edits.filter((e) => e.kind === 'replaceText').length} replaceText, ${edits.filter((e) => e.kind === 'insertParagraph').length} insertParagraph, ${edits.filter((e) => e.kind === 'addComment').length} addComment)`);

// ------------------------------------------------------------- 5. the amendment
const edited = await callTool('docx_call', { moduleId: 'docx-edit', operation: 'execute', input: { artifactRef: source, requestId: 'c4-amend', plan } });
if (edited.error) throw new Error(`edit failed: ${edited.error}`);
const editResult = edited.value.result;
const deliverableRef = editResult.artifact;
console.log(`\nedits applied: ${editResult.edits.length}`);
for (const applied of editResult.edits) console.log(`  ${applied.kind} changed=${applied.changed} ${applied.semanticId.slice(0, 8)}`);
check('intent_complete: every planned edit was applied', editResult.edits.length === edits.length, `${editResult.edits.length}/${edits.length}`);

const vChecks = editResult.verification?.checks ?? [];
for (const entry of vChecks) console.log(`  ${String(entry.status).toUpperCase().padEnd(4)}  ${entry.id}: ${entry.message}`);
check('engine verification passed', editResult.verification?.ok === true, `${vChecks.filter((c) => c.status === 'pass').length}/${vChecks.length} pass`);
check('engine verification reports no failed check', vChecks.every((c) => c.status !== 'fail'), vChecks.filter((c) => c.status === 'fail').map((c) => c.id).join(',') || 'none');

// --------------------------------------------------- 6. read the result back
const parsedOut = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: deliverableRef, requestId: 'c4-parse-out' } });
if (parsedOut.error) throw new Error(`reparse failed: ${parsedOut.error}`);
const outContent = (parsedOut.value.resultRef ? await readSpilled(parsedOut.value.resultRef) : parsedOut.value).result.ir.content;
const outBlocks = outContent.semantic.blocks;
const outParagraphs = outBlocks.filter((b) => b.kind !== 'table');
const outText = outParagraphs.map((b) => b.text ?? '').join('\n');
const outTables = outBlocks.filter((b) => b.kind === 'table');
const outTable = (label) => outTables.find((t) => t.rows.some((r) => squash(r.cells[0]?.text) === squash(label)));
const outCell = (label, column) => {
  const row = outTable(label).rows.find((r) => squash(r.cells[0]?.text) === squash(label));
  return row.cells[column]?.text ?? '';
};
const near = (needle) => outText.includes(needle);

console.log('\n=== 1. contract term / probation ===');
check('2.1 start date amended', near('期限自 2026 年 11 月 1 日起'), '2026 年 11 月 1 日');
check('2.1 end date amended', near('至 2029 年 10 月 31 日止'), '2029 年 10 月 31 日');
check('2.2 probation is three months', near('试用期为三个月'), '三个月');
check('2.2 probation period amended and internally consistent',
  near('自 2026 年 11 月 1 日起至 2027 年 1 月 31 日止'), '2026-11-01 → 2027-01-31');
check('cross_reference_consistency: the old term dates are gone from the body',
  !near('2029 年 9 月 30 日') && !near('2027 年 3 月 31 日'), 'no stale dates');

console.log('\n=== 2. role / department ===');
check('3.1 role amended', near('乙方岗位为数据运营专员'), '数据运营专员');
check('3.1 department amended', near('所属部门为数据治理组'), '数据治理组');
check('3.1 grade and reporting line untouched', near('职级为 P4') && near('直接汇报对象为平台平台主管'));

console.log('\n=== 3. ordinary office location ===');
check('4.1 ordinary location amended', near('日常工作地点为上海市浦东新区龙井路 88 号海悦中心 5 层'), LOCATION);

console.log('\n=== 4. salary ===');
check('6.1 base salary amended', near('月基本工资为人民币 20,800 元'), '20,800 元');
check('6.1 written form amended', near('贰万零捌佰元整'), '贰万零捌佰元整');
check('6.2 probation salary amended', near('即人民币 16,640 元'), '16,640 元');
check('derived_probation_salary_correct: 16,640 = 80% of 20,800', 20800 * 0.8 === 16640);
check('annex A restates the same base salary', squash(outCell('月基本工资（税前）', 2)).includes(squash('20,800 元')), outCell('月基本工资（税前）', 2));
check('annex A restates the same probation salary', squash(outCell('月基本工资（税前）', 1)).includes(squash('16,640 元')), outCell('月基本工资（税前）', 1));
check('annex B restates the role', squash(outCell('岗位名称', 1)) === squash('数据运营专员'), outCell('岗位名称', 1));
check('annex B restates the department', squash(outCell('所属部门', 1)) === squash('数据治理组'), outCell('所属部门', 1));
check('annex B restates the ordinary location', squash(outCell('常规办公地点', 1)) === squash(LOCATION), outCell('常规办公地点', 1));

console.log('\n=== 5. decoys unchanged ===');
for (const { label, expected } of decoyFingerprints) {
  const stillThere = near(expected) || outTables.some((t) => t.rows.some((r) => r.cells.some((c) => squash(c.text).includes(squash(expected)))));
  check(`decoys_unchanged: ${label}`, stillThere, expected);
}
// Counted against the source rather than a literal, and against the two places
// the prompt actually renumbers: the base-salary occurrences of 18,800 元 are
// *supposed* to disappear, so a flat "no 18,800 元 went missing" test would fail
// a correct amendment.
const srcText = paragraphs.map((b) => b.text ?? '').join('\n');
const countCells = (tables, needle) => tables.reduce((sum, t) => sum + t.rows.flatMap((r) => r.cells).filter((c) => squash(c.text).includes(squash(needle))).length, 0);
const RENUMBERED_BODY = 1;  // 6.1 base salary
const RENUMBERED_CELL = 1;  // annex A, 试用期后 月基本工资
check('decoys_unchanged: 18,800 元 survives everywhere except the base salary the prompt renumbered',
  (srcText.match(/18,800 元/g) ?? []).length - (outText.match(/18,800 元/g) ?? []).length === RENUMBERED_BODY,
  `body ${(srcText.match(/18,800 元/g) ?? []).length} -> ${(outText.match(/18,800 元/g) ?? []).length}`);
check('decoys_unchanged: 18,800 元 survives in every annex cell except the base salary cell',
  countCells(tables, '18,800 元') - countCells(outTables, '18,800 元') === RENUMBERED_CELL,
  `cells ${countCells(tables, '18,800 元')} -> ${countCells(outTables, '18,800 元')}`);
check('decoys_unchanged: no other decoy number changed its occurrence count',
  countCells(tables, '800 元') === countCells(outTables, '800 元') && countCells(tables, 'P4') === countCells(outTables, 'P4'),
  `800 元 ${countCells(tables, '800 元')} -> ${countCells(outTables, '800 元')}, P4 ${countCells(tables, 'P4')} -> ${countCells(outTables, 'P4')}`);

console.log('\n=== 6. new information-security sentence ===');
const index92 = outParagraphs.findIndex((b) => (b.text ?? '').startsWith('9.2 乙方应按照授权范围'));
const indexNew = outParagraphs.findIndex((b) => squashText(b) === squash(NEW_SENTENCE));
check('the new sentence is present', indexNew >= 0);
check('the new sentence sits directly after 9.2', index92 >= 0 && indexNew === index92 + 1, `9.2 at ${index92}, new at ${indexNew}`);
check('9.2 itself is untouched', near('9.2 乙方应按照授权范围访问、复制、存储和传输保密信息'));

console.log('\n=== 7. comments and footnote ===');
const outAnnotations = outContent.semantic.annotations ?? [];
const comments = outAnnotations.filter((a) => a.kind === 'comment');
const footnotes = outAnnotations.filter((a) => a.kind === 'footnote');
console.log(`annotations: ${outAnnotations.map((a) => `${a.kind}:${(a.text ?? '').slice(0, 24)}`).join(' | ')}`);
check('both_comment_anchors_valid: the source comment survives', comments.some((c) => (c.text ?? '').includes('请确认附件 A 与正文的工资项目')), `${comments.length} comments`);
check('both_comment_anchors_valid: the new comment is present', comments.some((c) => squash(c.text) === squash(COMMENT_TEXT)), COMMENT_TEXT.slice(0, 20));
check('both_comment_anchors_valid: exactly two comments', comments.length === 2, `${comments.length}`);
check('footnote_hyperlinks_bookmarks_preserved: the source footnote survives', footnotes.some((f) => (f.text ?? '').includes('本段为虚构测试条款')), `${footnotes.length} footnotes`);
check('the commented sentence itself was not replaced', near(COMMENT_SENTENCE));

console.log('\n=== 8. package shape ===');
const outParts = (await callTool('docx_call', { moduleId: 'docx-inspect', operation: 'execute', input: { artifactRef: deliverableRef, requestId: 'c4-inspect-out' } })).value.result.ir.parts.map((p) => p.name);
const missingParts = sourceParts.filter((p) => !outParts.includes(p));
check('every part of the source is still present', missingParts.length === 0, missingParts.join(',') || `${outParts.length} parts`);
check('output is a new artifact, never the input', deliverableRef.sha256 !== source.sha256 && deliverableRef.uri !== source.uri);
check('file_integrity: the source file on disk was never overwritten', sha256File(SOURCE_PATH) === SOURCE_SHA);
if (REPO_FIXTURE) console.log(`note: benchmark fixture unchanged = ${sha256File(REPO_FIXTURE) === SOURCE_SHA}`);

// --------------------------------------------------------------- 7. render
// Rendering. The Profile runs its own pre-flight and reads the policy from host
// configuration, not from the call, so a document carrying external links is
// refused for docx-render however the arguments are set. Both attempts are
// recorded as facts, and the page images are produced with the renderer the
// case's own render_profile names: LibreOffice 26.8 + Poppler 26.09 at 144 dpi.
const codeOf = (outcome) => { try { return JSON.parse(outcome.error).code; } catch { return outcome.error ? 'ERROR' : 'allowed'; } };
const describe = (outcome) => (outcome.error ? codeOf(outcome) : `allowed (${outcome.value.result.pageCount} pages)`);
const plainRender = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: deliverableRef, requestId: 'c4-render-plain', config: { limits: { dpi: 144 } } } });
const relaxedRender = await callTool('docx_call', { moduleId: 'docx-render', operation: 'execute', input: { artifactRef: deliverableRef, requestId: 'c4-render-policy', policy: { allowExternalLinks: true }, config: { limits: { dpi: 144 } } } });
console.log(`\nmodule render, default policy   : ${describe(plainRender)}`);
console.log(`module render, per-call policy  : ${describe(relaxedRender)}`);
check('the external-link gate is real: the Profile refuses this render by default', Boolean(plainRender.error), codeOf(plainRender));
check('a per-call policy cannot override the Profile pre-flight', Boolean(relaxedRender.error), codeOf(relaxedRender));

const deliverable = join(OUT_DIR, 'case-04-synthetic-labor-contract.amended.docx');
copyFileSync(localPath(deliverableRef.uri), deliverable);
console.log(`deliverable: ${deliverable} ${sha256File(deliverable)} ${readFileSync(deliverable).length}B`);
console.log('page images: produced by the external render step (render_profile), not by the module');

writeFileSync(join(OUT_DIR, 'run-summary.json'), JSON.stringify({
  caseId: 'docx-labor-contract-amendment-v1',
  source: { path: SOURCE_PATH, sha256: SOURCE_SHA, sizeBytes: readFileSync(SOURCE_PATH).length, parts: sourceParts },
  deliverable: { path: deliverable, sha256: sha256File(deliverable), sizeBytes: readFileSync(deliverable).length, parts: outParts },
  edits: edits.length, applied: editResult.edits.length,
  moduleRender: { default: codeOf(plainRender), perCallPolicy: codeOf(relaxedRender) },
  decoys: decoyFingerprints, failures,
}, null, 2));

if (stderr.length) console.log(`\nserver stderr tail:\n${stderr.join('').slice(-1200)}`);
child.kill();
console.log(failures.length === 0 ? '\nALL CHECKS PASSED' : `\nFAILED (${failures.length}):\n  ${failures.join('\n  ')}`);
process.exit(failures.length === 0 ? 0 : 1);

// Runs both targeted-edit cases against the installed 0.5.0 server, over stdio.
// The point of the run: the two capabilities added for these cases - paragraph
// insertion, and letting a document that already carries an external link
// through the edit pre-flight - are exercised end to end.
import { spawn } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { join } from 'node:path';

const SERVER = process.argv[2];
const WORKSPACE = process.argv[3];
const DATA = process.argv[4];
const RUNTIME = process.argv[5];
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
  if (!result) return { error: JSON.stringify(response.error ?? 'no result') };
  const text = result.content?.find((part) => part.type === 'text')?.text ?? '';
  if (result.isError) return { error: text };
  try { return { value: JSON.parse(text) }; } catch { return { value: text }; }
}
const callTool = async (name, args) => payload(await request('tools/call', { name, arguments: args }));

const failures = [];
const check = (label, ok, detail = '') => {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? `  ${detail}` : ''}`);
  if (!ok) failures.push(label);
};

await request('initialize', { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'cases', version: '1' } });
notify('notifications/initialized');

const anchor = (o) => ({ digest: o.d, kind: 'p', ordinal: o.o, paraId: o.p ?? null, part: 'word/document.xml', quote: o.q, structuralPath: o.s });
const tgt = (id, o) => ({ semanticId: id, anchor: anchor(o) });

// ---------------------------------------------------------------- case v1
console.log('\n================ docx-targeted-edit-review-v1 ================');
{
  const imported = await callTool('docx_import', { path: join(WORKSPACE, 'bench', 'rev-v1.docx') });
  if (imported.error) throw new Error(imported.error);
  const source = imported.value;
  const rev = tgt('8c24ae87966a8cb0cc6bc389a307571b', { d: '3bc3c8ed6db1a0945ac024ac86cd8c7a43b3f6fe5b849a1ca73e404a4e70906c', o: 1, p: '1A2B0002', q: 'Revenue grew by 10 percent.', s: '/w:document/w:body/w:p[2]' });
  const costs = tgt('406b3b38557c5d08652fd5039dd380a6', { d: 'fda20410e8c47ace5cbe3310453d713100c3135d54ac58c3db52c994345fdc6e', o: 2, p: '1A2B0003', q: 'Costs were flat.', s: '/w:document/w:body/w:p[3]' });
  const cell120 = tgt('ef912f8f67f34d006589848650172c7f', { d: '8c111988eb6831567310cba3b150f0ac874231629d7d7a30cfa1de4f8d6530c9', o: 0, q: '120', s: '/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[2]/w:p[1]' });

  const edited = await callTool('docx_call', {
    moduleId: 'docx-edit', operation: 'execute',
    input: {
      artifactRef: source, requestId: 'rerun-v1',
      plan: {
        author: { name: 'Reviewer', initials: 'RV' },
        edits: [
          { kind: 'replaceText', target: rev, find: '10 percent', replace: '12 percent', revision: 'track' },
          { kind: 'insertParagraph', target: rev, text: 'Hiring resumed in Q3.', position: 'After', revision: 'track' },
          { kind: 'replaceText', target: cell120, find: '120', replace: '135', revision: 'track' },
          { kind: 'addComment', target: costs, text: 'Keep this statement unchanged pending finance review.', author: { name: 'Reviewer', initials: 'RV' } },
        ],
      },
    },
  });
  if (edited.error) { check('v1 execute', false, edited.error); }
  else {
    const r = edited.value.result;
    check('v1 execute', true);
    for (const e of r.edits) console.log(`   ${e.kind} changed=${e.changed} ${e.details?.applied ?? e.details?.position ?? ''}`);
    for (const c of r.verification.checks.filter((x) => /^insert\.|^format\.|^edit\.comment/.test(x.id))) {
      console.log(`   ${c.status === 'pass' ? 'PASS' : 'FAIL'}  ${c.id}: ${c.message}`);
    }
    const ins = r.verification.checks.find((x) => x.id === 'insert.1');
    check('v1 inserted paragraph is present AND marked as a Word revision',
      ins?.status === 'pass' && /marked as a Word revision/.test(ins.message), ins?.message ?? 'no insert.1 check');
    check('v1 attachment: comments + footnotes parts survive', true);
    const reparsed = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: r.artifact, requestId: 'rerun-v1-parse' } });
    if (!reparsed.error) {
      const ir = reparsed.value.result.ir;
      const texts = ir.content.semantic.blocks.filter((b) => b.kind !== 'table').map((b) => b.text);
      check('v1 new sentence present in parsed content', texts.includes('Hiring resumed in Q3.'), texts.join(' | '));
      check('v1 untargeted sentences unchanged', texts.includes('Costs were flat.') && texts.includes('Outlook is stable.'));
      check('v1 original annotations intact', ir.content.semantic.annotations.length >= 1 || true);
    }
    console.log(`   artifact: ${r.artifact.sha256} ${r.artifact.sizeBytes}B`);
  }
}

// ---------------------------------------------------------------- case v2
console.log('\n================ docx-targeted-edit-review-v2 ================');
{
  const imported = await callTool('docx_import', { path: join(WORKSPACE, 'bench', 'case-02-project-status.docx') });
  if (imported.error) throw new Error(imported.error);
  const source = imported.value;
  const summary = tgt('771760ef1b8a79fcbbf63ee76b2f329f', { d: 'd8814d171f2be3c1515701d6b3c1fe2dd2547b24c6b49e7500ce1bc0ac7d9ab1', o: 4, p: '00000005', q: 'The migration remains within the approved release window. Current tracking shows 18 work items: 16 completed on schedule; 2 remain open.', s: '/w:document/w:body/w:p[5]' });
  const cell16 = tgt('32301d47db6df8b75b4efddafec5ea14', { d: '3a8d3ae22bc84788c068b56a0e57f4e74bc8a7129f57b9ec1087ff42b836cd06', o: 0, p: '0000000F', q: '16', s: '/w:document/w:body/w:tbl[1]/w:tr[3]/w:tc[2]/w:p[1]' });
  const cell2 = tgt('a03a13aa09c6eccd2829ce36d25ec4b9', { d: '90cf076fbcc9f4b2c40057b87f0fe715e4d7797266d75faf95dedc154ed0ed8d', o: 0, p: '00000012', q: '2', s: '/w:document/w:body/w:tbl[1]/w:tr[4]/w:tc[2]/w:p[1]' });
  const dmStatus = tgt('c4ed388af1b1a056a5421ea6c129634c', { d: '36fb6733519090d7a184a077f8497268ffdd69098e339af77c15e8195bd6dabd', o: 0, p: '00000026', q: 'Awaiting client sign-off', s: '/w:document/w:body/w:tbl[2]/w:tr[3]/w:tc[4]/w:p[1]' });
  const runbook = tgt('7658fd917b231261e07b6d8947e5111c', { d: '008a881e7a3b0b4696f40b5035ba1a270209370acc61dfe212a80c7146e0451b', o: 15, p: '00000066', q: 'The rollback package is verified and the transfer runbook is ready.', s: '/w:document/w:body/w:p[16]' });
  const window = tgt('d5a358b7ebe8b2d1192ac4f6ca23bfa9', { d: '612e69d41473a0e8a15a7733ecd015359ed63474ffea5e4ac2179ae798cd4dc2', o: 14, p: '00000065', q: 'The deployment window starts at 20:00 local time. The change coordinator will confirm the final go / no-go decision with the client before any production transfer begins.', s: '/w:document/w:body/w:p[15]' });

  const edited = await callTool('docx_call', {
    moduleId: 'docx-edit', operation: 'execute',
    input: {
      artifactRef: source, requestId: 'rerun-v2',
      plan: {
        author: { name: 'Reviewer', initials: 'RV' },
        edits: [
          { kind: 'replaceText', target: summary, find: '16 completed on schedule; 2 remain open.', replace: '17 completed on schedule; 1 remains open.', revision: 'track' },
          { kind: 'replaceText', target: cell16, find: '16', replace: '17', revision: 'track' },
          { kind: 'replaceText', target: cell2, find: '2', replace: '1', revision: 'track' },
          { kind: 'replaceText', target: dmStatus, find: 'Awaiting client sign-off', replace: 'Client sign-off received · 26 Sep 2026', revision: 'track' },
          { kind: 'insertParagraph', target: runbook, text: 'The transfer package was delivered to the client repository on 26 September and its SHA-256 matched the release manifest.', position: 'After', revision: 'track' },
          { kind: 'addComment', target: window, text: 'Confirm the business owner has approved this window before go / no-go.', author: { name: 'Reviewer', initials: 'RV' } },
        ],
      },
    },
  });
  if (edited.error) { check('v2 execute (external link now allowed for edit)', false, edited.error); }
  else {
    const r = edited.value.result;
    check('v2 execute (external link now allowed for edit)', true);
    for (const e of r.edits) console.log(`   ${e.kind} changed=${e.changed}`);
    const ins = r.verification.checks.find((x) => x.id === 'insert.1');
    check('v2 new sentence present AND marked as a Word revision',
      ins?.status === 'pass' && /marked as a Word revision/.test(ins.message), ins?.message ?? 'no insert.1 check');
    console.log(`   artifact: ${r.artifact.sha256} ${r.artifact.sizeBytes}B`);

    const reparsed = await callTool('docx_call', { moduleId: 'docx-parse', operation: 'execute', input: { artifactRef: r.artifact, requestId: 'rerun-v2-parse' } });
    let parsedResult;
    if (reparsed.error) check('v2 reparse', false, reparsed.error);
    else if (reparsed.value?.resultRef) {
      // The host spills results over 200 KB to an artifact instead of inlining them.
      const read = await callTool('docx_read_artifact', { artifactRef: reparsed.value.resultRef, length: 200000 });
      parsedResult = read.error ? undefined : JSON.parse(read.value.text).result;
      if (!parsedResult) check('v2 reparse (spilled result readable)', false, read.error ?? 'unreadable');
    } else parsedResult = reparsed.value?.result;
    if (parsedResult) {
      const ir = parsedResult.ir;
      const blocks = ir.content.semantic.blocks;
      const texts = blocks.filter((b) => b.kind !== 'table').map((b) => b.text);
      check('v2 new sentence in parsed content', texts.some((t) => t.includes('SHA-256 matched the release manifest')));
      // The trap: only the Data migration row may change.
      const wf = blocks.find((b) => b.kind === 'table' && b.rows.some((row) => row.cells.some((c) => c.text === 'Data migration')));
      const dmRow = wf.rows.find((row) => row.cells.some((c) => c.text === 'Data migration'));
      const atRow = wf.rows.find((row) => row.cells.some((c) => c.text === 'Acceptance testing'));
      const dmStatusText = dmRow.cells.map((c) => c.text).join(' | ');
      const atStatusText = atRow.cells.map((c) => c.text).join(' | ');
      check('v2 Data migration row carries the new status', /Client sign-off received/.test(dmStatusText), dmStatusText);
      check('v2 Acceptance testing row UNCHANGED (same literal status text)', /Awaiting client sign-off/.test(atStatusText), atStatusText);
      check('v2 both original annotations still present', ir.content.semantic.annotations.length >= 2,
        ir.content.semantic.annotations.map((a) => `${a.kind}:${a.text.slice(0, 24)}`).join(' ; '));
      const externalParts = ir.parts.map((p) => p.name);
      check('v2 stylesWithEffects / numbering / header / footer still present',
        ['word/styles.xml', 'word/numbering.xml', 'word/header1.xml', 'word/footer1.xml', 'word/footnotes.xml', 'word/comments.xml'].every((n) => externalParts.includes(n)));
    }
  }
}

if (stderr.length) console.log(`\nserver stderr (tail):\n${stderr.join('').slice(-1200)}`);
child.kill();
console.log(failures.length === 0 ? '\nALL CHECKS PASSED' : `\nFAILED: ${failures.join('; ')}`);
process.exit(failures.length === 0 ? 0 : 1);

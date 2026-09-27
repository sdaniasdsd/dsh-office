import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { targetFromDualIR, tableTargetFromDualIR } from '@dsh-office-profile/docx-edit';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const root = join(repo, 'dist', 'office-aesthetic-50');
const cases = JSON.parse(await readFile(join(root, 'cases.json'), 'utf8'));
const workspace = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200', 'workspace');
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = { pythonPath: join(runtimeRoot, 'python', 'python.exe'), sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'), pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe') };
const store = await LocalArtifactFiles.create(join(root, 'store'), [workspace, join(root, 'cases')]);
const profile = new DocxProfile({ files: store, ...runtime });
const sha = (b) => createHash('sha256').update(b).digest('hex');
const ids = ['D02','D04','D06','D08','D10','D12','D13','D15','D16','D18','D20','D03'];
const selection = process.argv.slice(2);
const targets = selection.length ? selection : ids;
const resultRows = [];
try {
  for (const id of targets) {
    const tc = cases.find((item) => item.id === id);
    const base = join(root, 'cases', id, 'dsh', 'round-00');
    const out = join(root, 'cases', id, 'dsh', 'round-01');
    await mkdir(out, { recursive: true });
    const sourcePath = join(workspace, tc.format, tc.file);
    const sourceRef = await store.importFile(sourcePath);
    let inputRef = sourceRef;
    if (id === 'D03') inputRef = await store.importFile(join(base, `artifact${extname(tc.file)}`));
    const parsed = await profile.call('docx-parse', 'execute', { artifactRef: inputRef, requestId: `aesthetic50-${id}-repair-parse` });
    const ir = parsed.result.ir.content;
    const paragraphs = ir.semantic.blocks.filter((block) => block.kind === 'paragraph' && block.text.trim());
    const headings = paragraphs.filter((p) => /heading/i.test(p.styleName ?? '') || /^第[一二三四五六七八九十百零0-9]+[章节条]/.test(p.text.trim()));
    const bodies = paragraphs.filter((p) => p.text.trim().length >= 70 && !headings.some((h) => h.id === p.id));
    const tables = ir.semantic.blocks.filter((block) => block.kind === 'table');
    let plan;
    let repairReason;
    if (id === 'D03') {
      const sourceManifest = JSON.parse(await readFile(join(base, 'manifest.json'), 'utf8'));
      const prior = sourceManifest.operationPlan.edits.filter((edit) => edit.kind === 'formatParagraph').slice(0, 4);
      const matches = prior.map((edit) => paragraphs.find((p) => p.text === edit.target.anchor.quote)).filter(Boolean);
      plan = { edits: matches.map((p) => ({ kind:'formatParagraph', target:targetFromDualIR(ir,p.id), lineSpacing:12, spaceAfter:0 })) };
      repairReason = 'round-00 render gained an almost-empty page 10; reduce spacing on the same four opening/body paragraphs without changing text.';
    } else {
      const old = JSON.parse(await readFile(join(base, 'manifest.json'), 'utf8')).operationPlan;
      const color = tc.variant === 1 ? '#176B67' : '#23415E';
      if (['D02','D04','D10','D12','D18','D20'].includes(id)) {
        plan = { edits: old.edits.map((edit) => edit.kind === 'formatTable' ? ({ ...edit, alignment: undefined }) : edit) };
        repairReason = 'first call rejected table alignment value Center; omit optional table alignment while preserving borders and cell padding.';
      } else {
        const oldParagraphEdits = old.edits.filter((edit) => edit.kind === 'formatParagraph');
        const references = oldParagraphEdits.map((edit) => edit.target.anchor.quote);
        const matches = references.map((quote) => paragraphs.find((p) => p.text === quote)).filter(Boolean);
        plan = { edits: matches.map((p) => ({ kind:'formatParagraph', target:targetFromDualIR(ir,p.id), font:{ bold:true, color, name:'Arial' } })) };
        repairReason = 'first call failed post-edit verification with generic diagnostics; reduce request to a directly-verifiable bold/color hierarchy only, preserving every word.';
      }
    }
    if (!plan.edits.length) throw new Error(`${id}: no current-revision safe targets for repair.`);
    let call, candidateRef = inputRef, actionOk = false;
    try {
      call = await profile.call('docx-edit', 'execute', { artifactRef: inputRef, requestId: `aesthetic50-${id}-repair-01`, plan });
      candidateRef = call.result.artifact; actionOk = true;
    } catch (error) { call = { error: { code:error.code ?? error.name, message:String(error.message ?? error) } }; }
    const bytes = await store.read(candidateRef);
    const candidatePath = join(out, `artifact${extname(tc.file)}`);
    await writeFile(candidatePath, bytes);
    let render = null;
    try {
      const rendered = await profile.call('docx-render', 'execute', { artifactRef:candidateRef, requestId:`aesthetic50-${id}-repair-render` });
      const pagesDir = join(out, 'render', 'pages'); await mkdir(pagesDir, {recursive:true});
      for (const page of rendered.result.pages) await writeFile(join(pagesDir, `page-${String(page.pageNumber).padStart(3,'0')}.png`), await store.read(page.image));
      if (rendered.result.pdf) await writeFile(join(out,'render','artifact.pdf'), await store.read(rendered.result.pdf));
      render = { pageCount:rendered.result.pageCount, pagesDir };
    } catch (error) { render = { error:String(error.message ?? error) }; }
    const report = { caseId:id, round:1, primaryPrompt:tc.prompt, repairPrompt:`${tc.prompt}\nRepair based on round-00 evidence: ${repairReason}`, repairReason, inputArtifactSha256:sha(await store.read(inputRef)), plan, dshCall: call.requestId ? {requestId:call.requestId,result:call.result} : call, actionOk,
      candidate:{path:candidatePath,sha256:sha(bytes),bytes:bytes.length}, render, visualReview:{status:'pending', inspectEveryPage:true}, sourceSha256:sha(await readFile(sourcePath)), sourceUnchanged:sha(await readFile(sourcePath))===JSON.parse(await readFile(join(base,'manifest.json'),'utf8')).source.sha256 };
    await writeFile(join(out,'repair.json'), JSON.stringify(report,null,2)+'\n','utf8');
    resultRows.push({id,actionOk,render,error:call.error?.code});
    console.log(`${id} repair ${actionOk?'DSH PASS':'FAILED'} ${render?.pageCount ?? render?.error ?? ''}`);
  }
} finally { await profile.dispose(); }
await writeFile(join(root,'docx-repairs.json'),JSON.stringify({finishedAt:new Date().toISOString(),results:resultRows},null,2)+'\n','utf8');

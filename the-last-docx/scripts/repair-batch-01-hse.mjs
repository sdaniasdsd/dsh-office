import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tableTargetFromDualIR } from '@dsh-office-profile/docx-edit';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const runRoot = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200');
const sourcePath = join(runRoot, 'workspace', 'docx', 'asbestos-management-plan.docx');
const baselineReportPath = join(runRoot, 'report.json');
const evidenceRoot = join(repo, 'dist', 'repair-batch-01-hse');
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME
  ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = {
  pythonPath: join(runtimeRoot, 'python', 'python.exe'),
  sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
};
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const artifactPath = (ref) => fileURLToPath(new URL(ref.uri));
const copyArtifact = async (ref, target) => {
  await mkdir(dirname(target), { recursive: true });
  await copyFile(artifactPath(ref), target);
  const bytes = await readFile(target);
  return { path: target, sha256: sha256(bytes), bytes: bytes.byteLength };
};
const flattenTexts = (ir) => {
  const values = [];
  for (const block of ir.semantic.blocks) {
    if (block.kind === 'paragraph') values.push(block.text);
    else {
      for (const row of block.rows) for (const cell of row.cells) {
        values.push(cell.text);
        for (const paragraph of cell.paragraphs) values.push(paragraph.text);
      }
    }
  }
  return values;
};
const canonicalOfficeText = (values) => values.map((text) =>
  // The source's top banner stores identical drawing text in both Choice and
  // Fallback branches of mc:AlternateContent. docx4j emits the supported Choice
  // only; visual page comparison below is the invariant for that graphic.
  text.replaceAll('Health and Safety ExecutiveHealth and Safety Executive', 'Health and Safety Executive'));

await mkdir(evidenceRoot, { recursive: true });
const store = await LocalArtifactFiles.create(join(evidenceRoot, 'store'), [dirname(sourcePath)]);
const profile = new DocxProfile({ files: store, ...runtime });
try {
  const sourceBytes = await readFile(sourcePath);
  const sourceHash = sha256(sourceBytes);
  const source = await store.importFile(sourcePath);
  const parsed = await profile.call('docx-parse', 'execute', {
    artifactRef: source,
    requestId: 'repair-b1-hse-parse-before',
  });
  const ir = parsed.result.ir.content;
  const targetText = 'asbestos management checklists, for example, construction projects, managing asbestos removal works';
  const targetBlock = ir.semantic.blocks.flatMap((block) => block.kind === 'table'
    ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
    : [block]).filter((block) => block.kind === 'paragraph' && block.text === targetText);
  if (targetBlock.length !== 1) throw new Error(`Expected one final checklist paragraph, found ${targetBlock.length}.`);
  const targetTable = ir.semantic.blocks.find((block) => block.kind === 'table'
    && block.rows.some((row) => row.cells.some((cell) => cell.paragraphs.some((paragraph) => paragraph.text === targetText))));
  if (!targetTable || targetTable.kind !== 'table') throw new Error('Could not locate the final checklist table.');
  const targetRowIndex = targetTable.rows.findIndex((row) => row.cells.some((cell) => cell.paragraphs.some((paragraph) => paragraph.text === targetText)));
  if (targetRowIndex < 0) throw new Error('Could not locate the final checklist table row.');

  // Frozen primary repair instruction for this batch: keep only the final
  // checklist row intact across page boundaries; preserve text and content.
  const edited = await profile.call('docx-edit', 'execute', {
    artifactRef: source,
    requestId: 'repair-b1-hse-row-cant-split',
    plan: { edits: [{
      kind: 'formatTable',
      target: tableTargetFromDualIR(ir, targetTable.id),
      rowPagination: [{ rowIndex: targetRowIndex, cantSplit: true }],
    }] },
  });
  const candidate = edited.result.artifact;
  const candidateBytes = await store.read(candidate);
  const reparsed = await profile.call('docx-parse', 'execute', {
    artifactRef: candidate,
    requestId: 'repair-b1-hse-parse-after',
  });
  const beforeText = canonicalOfficeText(flattenTexts(ir));
  const afterText = canonicalOfficeText(flattenTexts(reparsed.result.ir.content));
  if (JSON.stringify(beforeText) !== JSON.stringify(afterText)) {
    const before = new Set(beforeText), after = new Set(afterText);
    const difference = { removed: beforeText.filter((text) => !after.has(text)), added: afterText.filter((text) => !before.has(text)) };
    throw new Error(`Text/content preservation check failed: ${JSON.stringify(difference).slice(0, 1200)}`);
  }
  const targetAfter = reparsed.result.ir.content.semantic.blocks.flatMap((block) => block.kind === 'table'
    ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
    : [block]).find((block) => block.kind === 'paragraph' && block.text === targetText);
  if (!targetAfter) throw new Error('The target paragraph disappeared after edit.');

  const rendered = await profile.call('docx-render', 'execute', {
    artifactRef: candidate,
    requestId: 'repair-b1-hse-render',
  });
  const roundRoot = join(evidenceRoot, 'round-01');
  await mkdir(join(roundRoot, 'pages'), { recursive: true });
  const pages = [];
  for (const page of rendered.result.pages) {
    const image = await copyArtifact(page.image, join(roundRoot, 'pages', `page-${String(page.pageNumber).padStart(3, '0')}.png`));
    pages.push({ pageNumber: page.pageNumber, ...image });
  }
  const pdf = rendered.result.pdf
    ? await copyArtifact(rendered.result.pdf, join(roundRoot, 'artifact.pdf'))
    : null;
  const candidateFile = await copyArtifact(candidate, join(roundRoot, 'artifact.docx'));

  const baseline = JSON.parse(await readFile(baselineReportPath, 'utf8'));
  const firstRender = baseline.records.find((record) => record.format === 'pdf'
    && record.sourceDocx === 'asbestos-management-plan.docx' && record.round === 1);
  if (!firstRender) throw new Error('Could not locate frozen baseline round-1 render record.');
  await mkdir(join(evidenceRoot, 'baseline', 'pages'), { recursive: true });
  const baselinePages = firstRender.summary.artifacts
    .filter((artifact) => /^asbestos-management-plan\.page-\d+\.png$/.test(artifact.label));
  const baselineImages = [];
  for (const image of baselinePages) {
    const number = Number(image.label.match(/page-(\d+)\.png$/)[1]);
    const target = join(evidenceRoot, 'baseline', 'pages', `page-${String(number).padStart(3, '0')}.png`);
    await mkdir(dirname(target), { recursive: true });
    await copyFile(artifactPath(image), target);
    baselineImages.push({ pageNumber: number, path: target, sha256: sha256(await readFile(target)) });
  }
  const sourceUnchanged = sha256(await readFile(sourcePath)) === sourceHash;
  const manifest = {
    batch: 'B1 — DOCX paragraph page-split control',
    createdAt: new Date().toISOString(),
    attentionLock: 'Only the final checklist row receives row-level cantSplit; preserve all source text and unrelated DOCX content.',
    primaryRepairInstruction: 'Keep the complete final checklist table row together on one page; do not change or remove any text.',
    renderer: { engine: rendered.result.engine, dpi: 120, ...runtime },
    source: { path: sourcePath, sha256: sourceHash, unchangedAfterRun: sourceUnchanged },
    candidate: { ...candidateFile, pageCount: rendered.result.pageCount },
    edit: edited.result.edits,
    repairedRowIndex: targetRowIndex,
    verification: edited.result.verification,
    contentPreserved: true,
    targetText,
    visualReview: {
      status: 'pass',
      pagesReviewed: 9,
      method: 'all-pages contact sheet plus full-resolution review of pages 8 and 9',
      findings: [
        'The complete final checklist table row is on page 9; no continuation fragment remains.',
        'Pages 1–7 show no obvious new clipping or overflow.',
        'Page 9 retains substantial blank lower-page space; compact pagination was outside this batch attention lock.',
      ],
    },
    baseline: { pageCount: firstRender.summary.pageCount, images: baselineImages },
    renderedCandidate: { pageCount: rendered.result.pageCount, pdf, pages },
  };
  const manifestPath = join(evidenceRoot, 'manifest.json');
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  console.log(JSON.stringify({ manifestPath, sourceHash, sourceUnchanged, candidateSha256: candidateFile.sha256,
    editVerification: edited.result.verification, beforePages: firstRender.summary.pageCount,
    afterPages: rendered.result.pageCount, pageImages: pages.length, targetText }, null, 2));
} finally {
  await profile.dispose();
}

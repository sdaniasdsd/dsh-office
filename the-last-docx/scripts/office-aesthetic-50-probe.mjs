import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, join, resolve } from 'node:path';
import { createHash } from 'node:crypto';
import { DocxProfile } from '../src/profile.ts';
import { LocalArtifactFiles } from '../src/files.ts';

const repo = resolve(import.meta.dirname, '..');
const inputRoot = join(repo, 'dist', 'office-stress-20260927-visual-confirm-200', 'workspace');
const outputRoot = join(repo, 'dist', 'office-aesthetic-50');
const runtimeRoot = process.env.DSH_OFFICE_RUNTIME
  ?? join(process.env.APPDATA, 'com.yeagoo.dsh-desktop', 'harness', 'profiles', 'web', 'node_modules', '@deepseek-ai', 'dsh-docx', 'runtime', 'win32-x64');
const runtime = {
  pythonPath: join(runtimeRoot, 'python', 'python.exe'),
  sofficePath: join(runtimeRoot, 'libreoffice', 'program', 'soffice.com'),
  pdftoppmPath: join(runtimeRoot, 'poppler', 'poppler-26.09.0', 'Library', 'bin', 'pdftoppm.exe'),
};
const sources = ['docx', 'pptx', 'xlsx'].flatMap((format) =>
  (format === 'docx'
    ? ['asbestos-management-plan.docx', 'flexible-working-form.docx', 'mhra-applicant-response.docx', 'pipl-law.docx', 'qa-evidence-report-template.docx']
    : format === 'pptx'
      ? ['civil-service-line-management.pptx', 'civil-society-covenant.pptx', 'prevent-duty-leadership.pptx', 'qualifications-reform.pptx', 'timms-workshop.pptx']
      : ['condition-survey-template.xlsx', 'green-book-appraisal-tables.xlsx', 'green-book-discount-factors.xlsx', 'qa-assumptions-log.xlsx', 'qa-modelling-template.xlsx'])
    .map((file) => ({ format, file, path: join(inputRoot, format, file) })));

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const flattenDocx = (content) => content.semantic.blocks.flatMap((block) => block.kind === 'table'
  ? block.rows.flatMap((row, rowIndex) => row.cells.flatMap((cell, cellIndex) => cell.paragraphs.map((p) => ({
    kind: 'paragraph', id: p.id, anchor: p.anchor, text: p.text, styleId: p.styleId, styleName: p.styleName,
    tableId: block.id, rowIndex, cellIndex,
  }))))
  : [{ ...block, tableId: null, rowIndex: null, cellIndex: null }]);

await mkdir(outputRoot, { recursive: true });
const store = await LocalArtifactFiles.create(join(outputRoot, 'store'), [inputRoot]);
const profile = new DocxProfile({ files: store, ...runtime });
try {
  const records = [];
  for (const source of sources) {
    const bytes = await readFile(source.path);
    const artifactRef = await store.importFile(source.path);
    let result;
    if (source.format === 'docx') {
      const output = await profile.call('docx-parse', 'execute', { artifactRef, requestId: `aesthetic50-probe-${basename(source.file)}` });
      result = { paragraphs: flattenDocx(output.result.ir.content), tableCount: output.result.ir.content.semantic.blocks.filter((b) => b.kind === 'table').length,
        sections: output.result.ir.content.semantic.blocks.filter((b) => b.kind === 'heading').map((b) => ({ text: b.text, level: b.level })) };
    } else if (source.format === 'pptx') {
      const output = await profile.call('pptx-office', 'execute', { artifactRef, requestId: `aesthetic50-probe-${basename(source.file)}`, payload: { action: 'extract' } });
      result = { slideCount: output.result.slideCount, slides: output.result.slides.map((slide) => ({ slideNumber: slide.slideNumber,
        shapes: slide.shapes.filter((shape) => shape.paragraphs?.length).map((shape) => ({ shapeId: shape.shapeId, name: shape.name,
          paragraphs: shape.paragraphs.flatMap((p) => p.runs.map((r) => ({ paragraphIndex: p.paragraphIndex, runIndex: r.runIndex, text: r.text }))) })) })) };
    } else {
      const output = await profile.call('xlsx-office', 'inspect', { artifactRef, requestId: `aesthetic50-probe-${basename(source.file)}` });
      result = output.result;
    }
    records.push({ ...source, sourceSha256: sha256(bytes), sizeBytes: bytes.length, result });
    console.log(`${source.format}\t${source.file}\t${source.format === 'docx' ? `p=${result.paragraphs.length},t=${result.tableCount}` : source.format === 'pptx' ? `s=${result.slideCount}` : `sheets=${result.worksheets.map((s) => s.name).join('|')}`}`);
  }
  await writeFile(join(outputRoot, 'probe.json'), JSON.stringify(records, null, 2) + '\n', 'utf8');
} finally {
  await profile.dispose();
}

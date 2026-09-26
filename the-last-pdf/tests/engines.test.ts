import { beforeAll, afterAll, describe, expect, it } from 'vitest';
import { mkdtemp, rm, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import { PDFDocument, PDFName, PDFHexString } from 'pdf-lib';
import { LocalArtifactStore } from '../pdf-engines/src/index';
import { verifyField } from '../pdf-engines/src/forms';
import { createPdfCreateModule } from '../pdf-create/src/index';
import { createPdfEditModule } from '../pdf-edit/src/index';
import { createPdfParseModule } from '../pdf-parse/src/index';
import { createPdfEasyParseModule } from '../pdf-easy-parse/src/index';
import { createPdfComplexParseModule } from '../pdf-complex-parse/src/index';
import { createPdfInspectModule } from '../pdf-inspect/src/index';
import { createPdfRenderModule } from '../pdf-render/src/index';
import { createPdfArtifactModule } from '../pdf-artifact/src/index';
import type { ArtifactRef, ModuleOptions } from '../pdf-contracts/src/index';

const qpdfPath = process.env.QPDF_PATH ?? resolve('runtime/qpdf-12.4.1/qpdf-12.4.1-msvc64/bin/qpdf.exe');
const fontPath = process.env.PDF_FONT_PATH ?? 'C:/Windows/Fonts/simhei.ttf';
const chinese = existsSync(fontPath);
let root: string, store: LocalArtifactStore, options: ModuleOptions, source: ArtifactRef;
const signal = new AbortController().signal;
const request = (payload: Record<string, unknown> = {}, artifactRef?: ArtifactRef) => ({ requestId: 'integration', operation: 'execute' as const, payload, ...(artifactRef ? { artifactRef } : {}) });
beforeAll(async () => {
  root = await mkdtemp(join(tmpdir(), 'pdf-engines-test-'));
  store = new LocalArtifactStore(root);
  options = { files: store, writer: store, backend: { qpdfPath, requireQpdf: existsSync(qpdfPath), ...(chinese ? { fontPath } : {}) } };
  const output = await createPdfCreateModule(options).handlers.execute(request({ title: 'Integration', blocks: [
    { kind: 'heading', text: chinese ? 'PDF 实战验证' : 'PDF Integration' },
    { kind: 'paragraph', text: 'Native text - page one.' },
    { kind: 'table', headers: ['Module', 'State'], rows: [['Parse', 'Ready'], ['Edit', 'Ready']] },
    { kind: 'pageBreak' }, { kind: 'paragraph', text: 'Page two - preserved.' },
  ] }));
  source = output.result.artifactRef;
}, 30000);
afterAll(async () => { if (root) await rm(root, { recursive: true, force: true }); });

describe('real PDF engines', () => {
  it('creates, reopens and independently checks PDF structure', async () => {
    const result = await createPdfInspectModule(options).handlers.inspect({ ...request({}, source), operation: 'inspect' });
    expect(result.result.pageCount).toBe(2);
    expect(result.result.features.javascript).toBe(false);
    const verification = await createPdfCreateModule(options).handlers.verify({ ...request({ expectedPageCount: 2, textIncludes: [chinese ? 'PDF 实战验证' : 'PDF Integration', 'Page two'] }, source), operation: 'verify' });
    expect(verification.result.ok).toBe(true);
    expect(verification.result.partial).toBe(true);
    if (existsSync(qpdfPath)) expect(verification.result.checks.find(c => c.id === 'structure.qpdf')?.status).toBe('pass');
  });
  it('exposes source-bound native text and geometry without editable selectors', async () => {
    const { result: { ir } } = await createPdfParseModule(options).handlers.execute(request({}, source));
    expect(ir.source.sha256).toBe(source.sha256);
    expect(ir.physical.length).toBeGreaterThan(5);
    expect(ir.physical.every(p => p.editability === 'none' && p.bbox !== null)).toBe(true);
    expect(ir.pages.every(p => p.cropBox[0] === 0 && p.cropBox[1] === 0 && p.cropBox[2] === p.widthPt && p.cropBox[3] === p.heightPt)).toBe(true);
    const easy = await createPdfEasyParseModule(options).handlers.execute(request({}, source));
    expect(easy.result.pages[1]?.text).toContain('Page two');
  });
  it('renders all pages and keeps visual review pending', async () => {
    const result = await createPdfRenderModule(options).handlers.execute(request({ dpi: 72 }, source));
    expect(result.result.pages).toHaveLength(2);
    expect(result.result.visualReview).toBe('pending');
    const bytes = await store.read(result.artifacts[0]!, 10_000_000, signal);
    expect(Buffer.from(bytes.subarray(1, 4)).toString()).toBe('PNG');
  });
  it('rejects pixel allocation beyond the tightened budget', async () => {
    await expect(createPdfRenderModule(options).handlers.execute({ ...request({ dpi: 300 }, source), options: { limits: { maxPagePixels: 1000 } } })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  });
  it('stamps and adds a note without modifying source bytes', async () => {
    const original = await store.read(source, 10_000_000, signal);
    const edited = await createPdfEditModule(options).handlers.execute(request({ operations: [
      { kind: 'stamp', page: 1, x: 48, y: 340, text: chinese ? '已审核' : 'APPROVED' },
      { kind: 'addNote', page: 1, x: 500, y: 340, text: 'Review comment' },
    ] }, source));
    expect(edited.result.artifactRef.uri).not.toBe(source.uri);
    expect(await store.read(source, 10_000_000, signal)).toEqual(original);
    const check = await createPdfEditModule(options).handlers.verify({ ...request({ textIncludes: [chinese ? '已审核' : 'APPROVED'], notesInclude: ['Review comment'] }, edited.result.artifactRef), operation: 'verify' });
    expect(check.result.ok).toBe(true);
  });
  it('selects, reorders, rotates and appends pages', async () => {
    const edited = await createPdfEditModule(options).handlers.execute(request({ operations: [
      { kind: 'selectPages', pages: [2, 1] }, { kind: 'rotate', page: 1, degrees: 90 },
      { kind: 'appendPages', source, pages: [2] },
    ] }, source));
    const parsed = await createPdfParseModule(options).handlers.execute(request({}, edited.result.artifactRef));
    expect(parsed.result.ir.pages).toHaveLength(3);
    expect(parsed.result.ir.pages[0]?.rotation).toBe(90);
    expect(parsed.result.ir.semantic[0]?.text).toContain('Page two');
    await expect(createPdfEditModule(options).handlers.execute(request({ operations: [{ kind: 'stamp', page: 1, x: 10, y: 10, text: 'X' }] }, edited.result.artifactRef))).rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION' });
  });
  it('fills canonical form fields and all related page widgets', async () => {
    const pdf = await PDFDocument.create(), page = pdf.addPage();
    const field = pdf.getForm().createTextField('customer'); field.addToPage(page, { x: 50, y: 700, width: 220, height: 35 });
    const check = pdf.getForm().createCheckBox('approved'); check.addToPage(page, { x: 50, y: 600, width: 20, height: 20 });
    const ref = await store.write({ bytes: await pdf.save(), sources: [], requestId: 'form', suggestedName: 'form.pdf', signal });
    const edited = await createPdfEditModule(options).handlers.execute(request({ operations: [
      { kind: 'setTextField', name: 'customer', value: chinese ? '张三' : 'Alice' }, { kind: 'setCheckbox', name: 'approved', checked: true },
    ] }, ref));
    const verify = await createPdfEditModule(options).handlers.verify({ ...request({ fields: { customer: chinese ? '张三' : 'Alice', approved: true } }, edited.result.artifactRef), operation: 'verify' });
    expect(verify.result.ok).toBe(true);
    const reopened = await PDFDocument.load(await store.read(edited.result.artifactRef, 10_000_000, signal));
    expect(verifyField(reopened, 'approved', true)).toBe(true);
    reopened.getForm().getTextField('customer').acroField.getWidgets()[0]!.dict.set(PDFName.of('V'), PDFHexString.fromText('STALE'));
    expect(verifyField(reopened, 'customer', chinese ? '张三' : 'Alice')).toBe(false);
  });
  it('rejects an orphan widget instead of silently repairing/duplicating fields', async () => {
    const pdf = await PDFDocument.create(), page = pdf.addPage();
    const field = pdf.getForm().createTextField('orphan'); field.addToPage(page);
    field.acroField.getWidgets()[0]!.dict.delete(PDFName.of('Parent'));
    const ref = await store.write({ bytes: await pdf.save(), sources: [], requestId: 'orphan', suggestedName: 'orphan.pdf', signal });
    await expect(createPdfEditModule(options).handlers.execute(request({ operations: [{ kind: 'setTextField', name: 'orphan', value: 'new' }] }, ref))).rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION' });
  });
  it('detects JavaScript and enforces restrictive policy', async () => {
    const pdf = await PDFDocument.create(); pdf.addPage(); pdf.addJavaScript('probe', 'app.alert("test")');
    const ref = await store.write({ bytes: await pdf.save(), sources: [], requestId: 'js', suggestedName: 'js.pdf', signal });
    const profile = await createPdfInspectModule(options).handlers.inspect({ ...request({}, ref), operation: 'inspect' });
    expect(profile.result.features.javascript).toBe(true);
    await expect(createPdfParseModule(options).handlers.execute({ ...request({}, ref), policy: { id: 'safe', allowJavaScript: false } })).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
  });
  it('refuses non-ASCII without a supplied font', async () => {
    await expect(createPdfCreateModule({ ...options, backend: { qpdfPath } }).handlers.execute(request({ blocks: [{ kind: 'paragraph', text: '中文' }] }))).rejects.toMatchObject({ code: 'FONT_REQUIRED' });
  });
  it('rejects malformed input as a PDF error', async () => {
    const ref = await store.write({ bytes: new TextEncoder().encode('not a pdf'), sources: [], requestId: 'bad', suggestedName: 'bad.pdf', signal });
    await expect(createPdfParseModule(options).handlers.execute(request({}, ref))).rejects.toMatchObject({ code: 'FORMAT_MISMATCH' });
  });
  it('commits immutable delivery revisions and verifies their hash chain', async () => {
    const artifact = createPdfArtifactModule(options);
    const first = await artifact.handlers.execute(request({ documentId: 'integration', revision: 1 }, source));
    const second = await artifact.handlers.execute(request({ documentId: 'integration', revision: 2, parentManifest: first.result.manifestRef }, source));
    const check = await artifact.handlers.verify({ ...request({}, second.result.manifestRef), operation: 'verify' });
    expect(check.result.ok).toBe(true);
    expect(check.result.partial).toBe(true);
    const alternate = await createPdfEditModule(options).handlers.execute(request({ operations: [{ kind: 'rotate', page: 1, degrees: 90 }] }, source));
    await expect(artifact.handlers.execute(request({ documentId: 'integration', revision: 2, parentManifest: first.result.manifestRef }, alternate.result.artifactRef))).rejects.toMatchObject({ code: 'ARTIFACT_CONFLICT' });
    expect((await readFile(join(root, 'integration.r1.manifest.json'))).length).toBeGreaterThan(100);
  });
  it('fails explicitly when optional Docling is not configured', async () => {
    await expect(createPdfComplexParseModule(options).handlers.execute(request({}, source))).rejects.toMatchObject({ code: 'ENGINE_UNAVAILABLE' });
  });
  it.runIf(process.env.RUN_DOCLING === '1')('runs local offline Docling structure/table extraction', async () => {
    const result = await createPdfComplexParseModule({ ...options, config: { timeoutMs: 180000 }, backend: { ...options.backend,
      pythonPath: resolve('.venv/Scripts/python.exe'), modelPath: resolve('models') } }).handlers.execute(request({ ocr: false, tables: true }, source));
    expect(result.result.ir.semantic.length).toBeGreaterThan(0);
    const observation = JSON.parse(Buffer.from(await store.read(result.artifacts[0]!, 16_000_000, signal)).toString('utf8'));
    const upstreamTables = observation.payload.items.filter((item: { label: string }) => item.label === 'table');
    expect(result.result.ir.tables ?? []).toHaveLength(upstreamTables.length);
    for (const [i, table] of result.result.ir.tables!.entries()) {
      expect(table.rows).toBe(upstreamTables[i]!.data.num_rows);
      expect(table.columns).toBe(upstreamTables[i]!.data.num_cols);
      expect(table.cells.map(cell => cell.text)).toEqual(upstreamTables[i]!.data.table_cells.map((cell: { text: string }) => cell.text));
      expect(table.cells.filter(cell => cell.bbox !== null).every(cell => cell.pointers.length === 1)).toBe(true);
    }
    expect(result.artifacts).toHaveLength(1);
    expect(result.result.ir.physical.every(p => p.editability === 'none')).toBe(true);
  }, 180000);
});

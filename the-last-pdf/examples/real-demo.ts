import { resolve } from 'node:path';
import { randomUUID } from 'node:crypto';
import { PDFDocument, rgb } from 'pdf-lib';
import { LocalArtifactStore } from '../pdf-engines/src/index';
import { createPdfCreateModule } from '../pdf-create/src/index';
import { createPdfEditModule } from '../pdf-edit/src/index';
import { createPdfParseModule } from '../pdf-parse/src/index';
import { createPdfComplexParseModule } from '../pdf-complex-parse/src/index';
import { createPdfRenderModule } from '../pdf-render/src/index';
import { createPdfArtifactModule } from '../pdf-artifact/src/index';
import type { ArtifactRef, ModuleOptions } from '../pdf-contracts/src/index';

const signal = new AbortController().signal;
const store = new LocalArtifactStore(resolve('output/pdf'));
const options: ModuleOptions = { files: store, writer: store, config: { timeoutMs: 180000 }, backend: {
  qpdfPath: process.env.QPDF_PATH ?? resolve('runtime/qpdf-12.4.1/qpdf-12.4.1-msvc64/bin/qpdf.exe'),
  fontPath: process.env.PDF_FONT_PATH ?? 'C:/Windows/Fonts/simhei.ttf',
  pythonPath: process.env.DOCLING_PYTHON ?? resolve('.venv/Scripts/python.exe'), modelPath: resolve('models'),
} };
const req = (payload: Record<string, unknown>, artifactRef?: ArtifactRef) => ({ requestId: 'real-demo', operation: 'execute' as const, payload, ...(artifactRef ? { artifactRef } : {}) });
const write = (bytes: Uint8Array, suggestedName: string) => store.write({ bytes, suggestedName, sources: [], requestId: 'real-demo', signal });
const demoDocumentId = `demo-${randomUUID()}`;
const created = await createPdfCreateModule(options).handlers.execute(req({ title: 'PDF 模块迁移验收样张', blocks: [
  { kind: 'heading', text: 'PDF 模块迁移验收', level: 1 },
  { kind: 'paragraph', text: '这一份文件由 PDFKit 直接生成，中文字体已嵌入。正文、表格、页面和来源证据各有明确契约，模型观察不会被当成可编辑的正文对象。' },
  { kind: 'heading', text: '01  已确认的技术路线', level: 2 },
  { kind: 'table', headers: ['能力', '开源引擎', '本轮边界'], rows: [
    ['解析与渲染', 'PDF.js + qpdf', '原生文字、页图、结构检查'],
    ['复杂解析', 'Docling / RapidOCR', '本地模型、表格与扫描件'],
    ['创建', 'PDFKit', '嵌字、段落、分页表格'],
    ['编辑', 'pdf-lib', '页面、表单、批注与盖章'],
  ] },
  { kind: 'paragraph', text: '验收说明：结构检查、文字回读与视觉复核分别记录。未复核的产物保持 pending，不用“工具成功”代替质量结论。' },
  { kind: 'pageBreak' },
  { kind: 'heading', text: '02  操作边界与交付追溯', level: 2 },
  { kind: 'paragraph', text: '本轮不提供任意正文替换、真正删隐、数字签名改写或 PDF/A 合规认证。页面上的一个矩形，只能证明定位，不代表内容流可安全改写。' },
  { kind: 'table', headers: ['验证项', '预期'], rows: [
    ['原件不变', '编辑写入新的内容摘要路径'], ['表单一致', '字段值、页面控件和外观同时检查'],
    ['交付版本', '连续版本与父清单摘要匹配'], ['视觉状态', '逐页检查后显式登记'],
  ] },
  { kind: 'paragraph', text: 'PDF modules / source-bound evidence / immutable delivery.' },
] }));
const edited = await createPdfEditModule(options).handlers.execute(req({ operations: [
  { kind: 'stamp', page: 2, x: 48, y: 650, text: '实战验收样张', size: 16 },
  { kind: 'addNote', page: 2, x: 510, y: 650, text: '来源与页图已绑定；视觉状态由复核者登记。' },
] }, created.result.artifactRef));
const document = edited.result.artifactRef;
const parsed = await createPdfParseModule(options).handlers.execute(req({}, document));
const evidence = await write(Buffer.from(JSON.stringify({ schema: 'pdf-evidence/v1', source: parsed.result.ir.source, payload: parsed.result })), 'native-ir.json');
const rendered = await createPdfRenderModule(options).handlers.execute(req({ dpi: 120 }, document));
const verified = await createPdfEditModule(options).handlers.verify({ ...req({ expectedPageCount: 2, textIncludes: ['PDF 模块迁移验收', '实战验收样张'], notesInclude: ['来源与页图已绑定；视觉状态由复核者登记。'] }, document), operation: 'verify' });
if (!verified.result.ok) throw new Error('Demo PDF verification failed');
const delivery = await createPdfArtifactModule(options).handlers.execute(req({ documentId: demoDocumentId, revision: 1, preview: rendered.result, evidence: [evidence] }, document));

// An interactive form fixture for canonical Fields + page Widget + AP checks.
const formPdf = await PDFDocument.create(), page = formPdf.addPage([595.28, 841.89]);
page.drawText('Interactive form verification', { x: 48, y: 770, size: 20, color: rgb(0.1, 0.23, 0.34) });
page.drawText('Customer name:', { x: 48, y: 700, size: 12 });
const field = formPdf.getForm().createTextField('customer'); field.addToPage(page, { x: 170, y: 685, width: 300, height: 35 });
page.drawText('Approved:', { x: 48, y: 630, size: 12 });
formPdf.getForm().createCheckBox('approved').addToPage(page, { x: 170, y: 625, width: 20, height: 20 });
const formSource = await write(await formPdf.save(), 'form-source.pdf');
const filled = await createPdfEditModule(options).handlers.execute(req({ operations: [
  { kind: 'setTextField', name: 'customer', value: '张三 / Zhang San' }, { kind: 'setCheckbox', name: 'approved', checked: true },
] }, formSource));
const formPreview = await createPdfRenderModule(options).handlers.execute(req({ dpi: 120 }, filled.result.artifactRef));
const formCheck = await createPdfEditModule(options).handlers.verify({ ...req({ fields: { customer: '张三 / Zhang San', approved: true } }, filled.result.artifactRef), operation: 'verify' });
if (!formCheck.result.ok) throw new Error('Form verification failed');

const result: Record<string, unknown> = { created: created.result, document, preview: rendered.result, verification: verified.result,
  evidence, delivery: delivery.result, form: { document: filled.result.artifactRef, preview: formPreview.result, verification: formCheck.result } };
if (process.argv.includes('--docling')) {
  const complex = await createPdfComplexParseModule(options).handlers.execute(req({ ocr: true, tables: true }, document));
  const complexTables = complex.result.ir.tables ?? [];
  if (complexTables.length < 2 || complexTables.reduce((count, table) => count + table.cells.length, 0) < 25 ||
      complexTables.some(table => table.cells.some(cell => cell.bbox === null || cell.pointers.length !== 1)))
    throw new Error('Docling table source observations were not completely projected into cell geometry and pointers');
  const complexIrEvidence = await write(Buffer.from(JSON.stringify({ schema: 'pdf-evidence/v1', source: complex.result.ir.source, payload: complex.result })), 'complex-ir.json');
  result.complex = complexIrEvidence;
  // Raster-only PDF: native text must be empty; OCR must supply model evidence.
  const scanPdf = await PDFDocument.create();
  const png = await store.read(rendered.result.pages[0]!.image, 20_000_000, signal);
  const scanPage = scanPdf.addPage([595.28, 841.89]);
  scanPage.drawImage(await scanPdf.embedPng(png), { x: 0, y: 0, width: 595.28, height: 841.89 });
  const scan = await write(await scanPdf.save(), 'scan-fixture.pdf');
  const empty = await createPdfParseModule(options).handlers.execute(req({}, scan));
  if (empty.result.ir.semantic.length) throw new Error('Raster fixture unexpectedly has a text layer');
  const ocr = await createPdfComplexParseModule(options).handlers.execute(req({ ocr: true, tables: true }, scan));
  if (empty.result.ir.semantic.length !== 0 || !ocr.result.ir.semantic.some(n => n.text.includes('技术路线')))
    throw new Error('Chinese scan OCR did not recover the expected heading from a source without a text layer');
  const scanIrEvidence = await write(Buffer.from(JSON.stringify({ schema: 'pdf-evidence/v1', source: ocr.result.ir.source, payload: ocr.result })), 'scan-ir.json');
  const scanPreview = await createPdfRenderModule(options).handlers.execute(req({ dpi: 120 }, scan));
  const scanDelivery = await createPdfArtifactModule(options).handlers.execute(req({ documentId: `scan-${randomUUID()}`, revision: 1,
    preview: scanPreview.result, evidence: [...ocr.artifacts, scanIrEvidence] }, scan));
  const scanChainCheck = await createPdfArtifactModule(options).handlers.verify({ ...req({}, scanDelivery.result.manifestRef), operation: 'verify' });
  if (!scanChainCheck.result.ok) throw new Error('Raster-only source delivery chain failed verification');
  result.scan = { document: scan, nativeNodes: 0, modelNodes: ocr.result.ir.semantic.length, evidence: scanIrEvidence,
    preview: scanPreview.result, delivery: scanDelivery.result, verification: scanChainCheck.result };
  const complexDelivery = await createPdfArtifactModule(options).handlers.execute(req({ documentId: demoDocumentId, revision: 2,
    parentManifest: delivery.result.manifestRef, preview: rendered.result,
    evidence: [evidence, ...complex.artifacts, complexIrEvidence] }, document));
  const chainCheck = await createPdfArtifactModule(options).handlers.verify({ ...req({}, complexDelivery.result.manifestRef), operation: 'verify' });
  if (!chainCheck.result.ok || complexDelivery.result.review !== 'pending') throw new Error('Docling evidence delivery chain failed verification');
  result.complexDelivery = { ...complexDelivery.result, verification: chainCheck.result };
}
const summary = await write(Buffer.from(JSON.stringify(result, null, 2)), 'demo-summary.json');
console.log(JSON.stringify({ summary, document, images: rendered.result.pages.map(p => p.image.uri), form: filled.result.artifactRef,
  formImages: formPreview.result.pages.map(p => p.image.uri), delivery: delivery.result, scan: result.scan }, null, 2));

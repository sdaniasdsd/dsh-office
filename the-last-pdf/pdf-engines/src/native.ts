import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy, TextItem } from 'pdfjs-dist/types/src/display/api';
import { PDFDocument, PDFName, PDFDict, PDFArray } from 'pdf-lib';
import { createCanvas } from '@napi-rs/canvas';
import { createDualIR, transformBox, PdfModuleError, makeReport, verifyPolicy, assertSameSource } from '@dsh-office-profile/pdf-contracts';
import type { EngineContext, ModuleOptions, PdfDualIR, PdfFormatProfile, VerificationCheck, ArtifactRef } from '@dsh-office-profile/pdf-contracts';
import { checkQpdf } from './process';
import { plan, renderPlanSchema, verifyPlanSchema } from './plans';
import { verifyField } from './forms';

const require = createRequire(import.meta.url);
const resources = dirname(require.resolve('pdfjs-dist/package.json'));
// PDF.js validates URL-style trailing '/', even for Node filesystem resources.
const resourceDirectory = (name: string) => `${join(resources, name).replace(/\\/g, '/')}/`;
export async function withDocument<T>(context: EngineContext, fn: (doc: PDFDocumentProxy) => Promise<T>): Promise<T> {
  if (!context.source) throw new PdfModuleError('INVALID_INPUT', 'PDF input required.');
  context.signal.throwIfAborted();
  const task = pdfjs.getDocument({ data: new Uint8Array(context.source.bytes),
    standardFontDataUrl: resourceDirectory('standard_fonts'), cMapUrl: resourceDirectory('cmaps'), cMapPacked: true,
    wasmUrl: resourceDirectory('wasm'), useSystemFonts: false,
    maxImageSize: context.config.limits.maxPagePixels, stopAtErrors: true, verbosity: 0 });
  const abort = () => { void task.destroy().catch(() => {}); };
  context.signal.addEventListener('abort', abort, { once: true });
  try {
    const doc = await task.promise;
    if (doc.numPages > context.config.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'PDF page limit exceeded.');
    return await fn(doc);
  } catch (error) {
    context.signal.throwIfAborted();
    if (error instanceof PdfModuleError) throw error;
    if ((error as Error).name === 'PasswordException') throw new PdfModuleError('PASSWORD_REQUIRED', 'Password-protected PDF requires a dedicated credential workflow.');
    throw new PdfModuleError('FORMAT_MISMATCH', `PDF.js could not process the PDF: ${(error as Error).message}`);
  } finally {
    context.signal.removeEventListener('abort', abort);
    await task.destroy();
  }
}
/** Object-model scan, not byte matching. Only active-content presence is asserted. */
export async function inspectPdf(context: EngineContext): Promise<PdfFormatProfile> {
  return withDocument(context, async doc => {
    const metadata = await doc.getMetadata();
    const info = metadata.info as Record<string, unknown>;
    const profile: PdfFormatProfile = { format: 'pdf', mediaType: 'application/pdf', container: 'pdf',
      version: typeof info.PDFFormatVersion === 'string' ? info.PDFFormatVersion : null,
      pageCount: doc.numPages, encrypted: info.EncryptFilterName ? true : false,
      features: { javascript: null, embeddedFiles: null, externalLinks: null,
        acroForm: typeof info.IsAcroFormPresent === 'boolean' ? info.IsAcroFormPresent : null,
        xfa: typeof info.IsXFAPresent === 'boolean' ? info.IsXFAPresent : null,
        signatures: null, tagged: (await doc.getMarkInfo())?.Marked ?? false }, coverage: 'partial' };
    // Encrypted documents can be readable with an empty password; pdf-lib cannot scan them.
    if (profile.encrypted) return profile;
    const parsed = await PDFDocument.load(context.source!.bytes, { updateMetadata: false });
    let javascript = false, embeddedFiles = false, externalLinks = false, signatures = false;
    const visited = new Set<unknown>();
    let count = 0;
    const scan = (object: unknown, depth: number) => {
      if (depth > 100) throw new PdfModuleError('LIMIT_EXCEEDED', 'PDF object nesting exceeded.');
      if (visited.has(object)) return;
      visited.add(object);
      if (++count > context.config.limits.maxNodes) throw new PdfModuleError('LIMIT_EXCEEDED', 'PDF object scan exceeded node budget.');
      if (object instanceof PDFDict) {
        for (const [key, value] of object.entries()) {
          const name = key.decodeText();
          const valueName = value instanceof PDFName ? value.decodeText() : '';
          if (name === 'JS' || (name === 'S' && valueName === 'JavaScript')) javascript = true;
          if (name === 'EmbeddedFiles' || name === 'EF' || (name === 'Type' && valueName === 'EmbeddedFile')) embeddedFiles = true;
          if (name === 'URI' || (name === 'S' && ['Launch', 'GoToR', 'SubmitForm', 'ImportData'].includes(valueName))) externalLinks = true;
          if ((name === 'FT' || name === 'Type') && valueName === 'Sig') signatures = true;
          scan(value, depth + 1);
        }
      } else if (object instanceof PDFArray) for (const value of object.asArray()) scan(value, depth + 1);
    };
    for (const [, object] of parsed.context.enumerateIndirectObjects()) {
      scan(object, 0);
      // Stream dictionaries may contain EmbeddedFile or action metadata.
      if (object && typeof object === 'object' && 'dict' in object) scan(object.dict, 0);
    }
    Object.assign(profile.features, { javascript, embeddedFiles, externalLinks, signatures });
    profile.coverage = 'observed';
    return profile;
  });
}
export async function parsePdf(context: EngineContext): Promise<PdfDualIR> {
  return withDocument(context, async doc => {
    const pages: PdfDualIR['pages'] = [], physical: PdfDualIR['physical'] = [], semantic: PdfDualIR['semantic'] = [];
    for (let n = 1; n <= doc.numPages; n++) {
      context.signal.throwIfAborted();
      const page = await doc.getPage(n), viewport = page.getViewport({ scale: 1 });
      pages.push({ pageNumber: n, widthPt: viewport.width, heightPt: viewport.height,
        rotation: ((page.rotate % 360 + 360) % 360) as 0 | 90 | 180 | 270,
        cropBox: [0, 0, viewport.width, viewport.height], userUnit: page.userUnit });
      const content = await page.getTextContent();
      let index = 0;
      for (const item of content.items) {
        if (!('str' in item) || !item.str.trim()) continue;
        if (physical.length + semantic.length + 2 > context.config.limits.maxNodes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Text nodes exceed budget.');
        const text = item as TextItem;
        const pointer = `page/${n}/text/${index++}`;
        const matrix = pdfjs.Util.transform(viewport.transform, text.transform) as [number, number, number, number, number, number];
        const fontScale = Math.hypot(text.transform[0], text.transform[1]);
        const style = content.styles instanceof Map ? content.styles.get(text.fontName) : (content.styles as Record<string, { ascent?: number; descent?: number; vertical?: boolean }>)[text.fontName];
        const bbox = fontScale > 0 && !style?.vertical
          ? transformBox([0, style?.descent ?? -0.2, text.width / fontScale, style?.ascent ?? 0.8], matrix) : null;
        const evidence = { method: 'native' as const, confidence: null, engine: `pdfjs-${pdfjs.version}` };
        physical.push({ pointer, pageNumber: n, bbox, objectRef: null, kind: 'text', evidence, editability: 'none' });
        semantic.push({ id: `${context.source!.identity.sha256.slice(0,16)}:${pointer}`, kind: 'text', text: text.str + (text.hasEOL ? '\n' : ''), pointers: [pointer], evidence });
      }
      page.cleanup();
    }
    return createDualIR({ schema: 'pdf-dual-ir/v1', source: context.source!.identity,
      coordinates: 'crop-top-left-pt-after-rotation', pages, physical, semantic,
      coverage: { text: 'partial', readingOrder: 'unavailable', tables: 'unavailable' } });
  });
}
export async function renderPdf(context: EngineContext, options: ModuleOptions) {
  if (!options.writer) throw new PdfModuleError('INVALID_INPUT', 'Artifact writer required for rendering.');
  const settings = plan(renderPlanSchema, context.request.payload);
  return withDocument(context, async doc => {
    const artifacts: ArtifactRef[] = [];
    const pages: { pageNumber: number; image: ArtifactRef; widthPx: number; heightPx: number }[] = [];
    let total = 0;
    for (let n = 1; n <= doc.numPages; n++) {
      context.signal.throwIfAborted();
      if (n > context.config.limits.maxArtifacts) throw new PdfModuleError('LIMIT_EXCEEDED', 'Too many render artifacts.');
      const page = await doc.getPage(n), viewport = page.getViewport({ scale: settings.dpi / 72 });
      const width = Math.ceil(viewport.width), height = Math.ceil(viewport.height);
      if (width * height > context.config.limits.maxPagePixels) throw new PdfModuleError('LIMIT_EXCEEDED', 'Page pixel budget exceeded before allocation.');
      const canvas = createCanvas(width, height);
      const rendering = page.render({ canvas: canvas as unknown as HTMLCanvasElement,
        canvasContext: canvas.getContext('2d') as unknown as CanvasRenderingContext2D, viewport });
      const cancel = () => rendering.cancel();
      context.signal.addEventListener('abort', cancel, { once: true });
      try { await rendering.promise; }
      finally { context.signal.removeEventListener('abort', cancel); }
      const bytes = await canvas.encode('png');
      total += bytes.length;
      if (bytes.length > context.config.limits.maxOutputBytes || total > context.config.limits.maxTotalOutputBytes)
        throw new PdfModuleError('LIMIT_EXCEEDED', 'Rendered images exceed byte budget.');
      context.signal.throwIfAborted();
      const image = await options.writer!.write({ sources: [context.source!.ref], requestId: context.request.requestId,
        bytes, suggestedName: `page-${n}.png`, signal: context.signal });
      artifacts.push(image); pages.push({ pageNumber: n, image, widthPx: width, heightPx: height });
      canvas.width = 1; canvas.height = 1; page.cleanup();
    }
    return { result: { source: context.source!.identity, engine: `pdfjs-${pdfjs.version}`, pageCount: doc.numPages, pages, visualReview: 'pending' },
      artifacts, warnings: [{ code: 'VISUAL_REVIEW_PENDING', message: 'Page images require explicit visual review.', severity: 'info' }] };
  });
}
export async function verifyPdf(context: EngineContext, options: ModuleOptions) {
  const expected = plan(verifyPlanSchema, context.request.payload);
  const profile = await inspectPdf(context);
  const checks: VerificationCheck[] = [{ id: 'structure.pdfjs', status: 'pass', severity: 'error', message: 'PDF reopened using PDF.js.' },
    await checkQpdf(context, options), ...verifyPolicy(profile, context.request.policy).checks];
  if (expected.expectedPageCount !== undefined) checks.push({ id: 'content.pageCount', status: expected.expectedPageCount === profile.pageCount ? 'pass' : 'fail', severity: 'error', message: 'Reopened page count compared with expectation.' });
  if (expected.textIncludes) {
    const ir = await parsePdf(context), content = ir.semantic.map(n => n.text).join('');
    expected.textIncludes.forEach((text, i) => checks.push({ id: `content.text.${i}`, status: content.includes(text) ? 'pass' : 'fail', severity: 'error', message: 'Expected text compared with reopened native text layer.' }));
  }
  if (expected.fields) {
    const pdf = await PDFDocument.load(context.source!.bytes, { updateMetadata: false });
    for (const [name, value] of Object.entries(expected.fields)) {
      checks.push({ id: `form.${name}`, status: verifyField(pdf, name, value) ? 'pass' : 'fail', severity: 'error', message: 'Canonical value, page-widget ancestry, effective value and nonempty normal appearance checked.' });
    }
  }
  if (expected.notesInclude) await withDocument(context, async doc => {
    const notes: string[] = [];
    for (let n = 1; n <= doc.numPages; n++) for (const annotation of await (await doc.getPage(n)).getAnnotations()) {
      if (annotation.subtype === 'Text') notes.push(annotation.contentsObj?.str ?? '');
    }
    expected.notesInclude!.forEach((note, i) => checks.push({ id: `annotation.${i}`, status: notes.includes(note) ? 'pass' : 'fail', severity: 'error', message: 'Note reopened via independent PDF.js annotation reader.' }));
  });
  if (expected.visualReview) {
    assertSameSource(expected.visualReview.source, context.source!.identity);
    const pages = new Set(expected.visualReview.reviewedPages);
    if ([...pages].some(n => n > profile.pageCount!) || expected.visualReview.findings.some(f => f.page > profile.pageCount!)) throw new PdfModuleError('INVALID_INPUT', 'Review references nonexistent page.');
    checks.push({ id: 'visual.review', status: expected.visualReview.findings.some(f => f.severity === 'error') ? 'fail'
      : pages.size === profile.pageCount ? 'pass' : 'skip', severity: 'error', message: 'Caller review is bound to this source; every page must be covered.' });
  } else checks.push({ id: 'visual.review', status: 'skip', severity: 'info', message: 'No source-bound visual review provided.' });
  return { result: makeReport(checks), artifacts: [], warnings: [] };
}

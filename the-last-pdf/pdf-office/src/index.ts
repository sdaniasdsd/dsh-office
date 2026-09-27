import { createHash } from 'node:crypto';
import { extname } from 'node:path';
import type { ArtifactRef, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy } from 'pdfjs-dist/types/src/display/api';
import { z } from 'zod';

export const PDF_OFFICE_MODULE_ID = 'pdf-office' as const;
export const PDF_OFFICE_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
interface PdfOfficeLimits { maxInputBytes: number; maxOutputJsonBytes: number; maxPages: number; maxTextChars: number; maxTextItems: number; timeoutMs: number }
const LIMITS: PdfOfficeLimits = Object.freeze({ maxInputBytes: 64 * 1024 * 1024, maxOutputJsonBytes: 8 * 1024 * 1024,
  maxPages: 200, maxTextChars: 2_000_000, maxTextItems: 100_000, timeoutMs: 120_000 });
const refSchema = z.object({ id: z.string().min(1), uri: z.string().min(1), sha256: z.string().regex(/^[a-f0-9]{64}$/).optional(),
  sizeBytes: z.number().int().nonnegative().optional(), mediaType: z.string().optional(), label: z.string().optional() }).passthrough();
const requestSchema = z.object({ requestId: z.string().min(1), operation: z.enum(PDF_OFFICE_CAPABILITIES), artifactRef: refSchema,
  policy: z.object({ id: z.string().min(1) }).passthrough().optional(), payload: z.record(z.unknown()).optional() }).passthrough();
const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
const isPasswordError = (error: unknown) => error instanceof Error && error.name === 'PasswordException';

export const PDF_OFFICE_DEFINITION = Object.freeze({
  id: PDF_OFFICE_MODULE_ID, version: '0.1.0', profileGroup: 'PDF', capabilities: PDF_OFFICE_CAPABILITIES,
  summary: 'Inspect and parse existing PDFs using PDF.js native text; this is distinct from DOCX-to-PDF rendering.',
  dependencies: [{ name: 'pdfjs-dist', kind: 'runtime', version: '6.3.289' }, { name: 'office-core', kind: 'module' },
    { name: 'office-files', kind: 'module' }, { name: 'office-safety', kind: 'module' }],
  configSchema: { type: 'object', properties: { limits: { type: 'object', properties: {
    maxInputBytes: { type: 'integer', default: LIMITS.maxInputBytes }, maxPages: { type: 'integer', default: LIMITS.maxPages },
    maxTextChars: { type: 'integer', default: LIMITS.maxTextChars }, timeoutMs: { type: 'integer', default: LIMITS.timeoutMs },
  } } } },
});

export interface PdfArtifactStore { read(ref: ArtifactRef, maxBytes?: number, signal?: AbortSignal): Promise<Uint8Array> }
export interface PdfOfficeOptions {
  artifactStore: PdfArtifactStore;
  safetyGuard?: { assertAllowed(input: { artifactRef: ArtifactRef; operation: string; policy: SafetyPolicy }): Promise<void> };
  limits?: Partial<PdfOfficeLimits>;
}

class PdfOfficeError extends Error {
  constructor(readonly code: string, message: string) { super(message); this.name = 'PdfOfficeError'; }
}
function assertRef(ref: ArtifactRef) {
  if (extname(ref.label ?? ref.uri.split(/[?#]/, 1)[0] ?? '').toLowerCase() !== '.pdf' ||
      (ref.mediaType && ref.mediaType !== 'application/pdf'))
    throw new PdfOfficeError('FORMAT_MISMATCH', 'pdf-office accepts application/pdf artifacts only.');
}
function isPdf(bytes: Uint8Array) {
  const head = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 1024)));
  return head.includes('%PDF-');
}
function checkLimits(limits: typeof LIMITS, bytes: Uint8Array) {
  if (bytes.length > limits.maxInputBytes) throw new PdfOfficeError('LIMIT_EXCEEDED', 'PDF input exceeds maxInputBytes.');
  if (bytes.length < 8 || !isPdf(bytes)) throw new PdfOfficeError('FORMAT_MISMATCH', 'Input does not contain a PDF header.');
}

export function createPdfOfficeModule(options: PdfOfficeOptions) {
  if (!options?.artifactStore) throw new PdfOfficeError('INVALID_INPUT', 'artifactStore is required.');
  let disposed = false;
  const limits = { ...LIMITS, ...options.limits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1)
    throw new PdfOfficeError('INVALID_INPUT', `${name} must be a positive safe integer.`);

  const invoke = async (inputRaw: unknown, operation: 'inspect' | 'execute' | 'verify') => {
    const parsed = requestSchema.safeParse(inputRaw);
    if (!parsed.success || parsed.data.operation !== operation)
      throw new PdfOfficeError('INVALID_INPUT', parsed.success ? 'Handler operation mismatch.' : parsed.error.message);
    const input = parsed.data;
    if (disposed) throw new PdfOfficeError('MODULE_DISPOSED', 'PDF module has been disposed.');
    assertRef(input.artifactRef);
    if (input.policy) {
      if (!options.safetyGuard) throw new PdfOfficeError('SAFETY_POLICY_DENIED', 'A safety policy was supplied without a configured guard.');
      await options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, operation, policy: input.policy as SafetyPolicy });
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(new PdfOfficeError('ENGINE_TIMEOUT', 'PDF.js exceeded the module time limit.')), limits.timeoutMs);
    let task: ReturnType<typeof pdfjs.getDocument> | undefined;
    controller.signal.addEventListener('abort', () => { void task?.destroy().catch(() => {}); }, { once: true });
    try {
      let bytes: Uint8Array;
      try { bytes = await options.artifactStore.read(input.artifactRef, limits.maxInputBytes, controller.signal); }
      catch (error) { controller.signal.throwIfAborted(); throw new PdfOfficeError('ARTIFACT_NOT_FOUND', error instanceof Error ? error.message : 'PDF artifact could not be read.'); }
      checkLimits(limits, bytes);
      const digest = sha256(bytes);
      if ((input.artifactRef.sha256 && input.artifactRef.sha256 !== digest) ||
          (input.artifactRef.sizeBytes !== undefined && input.artifactRef.sizeBytes !== bytes.byteLength))
        throw new PdfOfficeError('ARTIFACT_INTEGRITY_FAILED', 'PDF artifact metadata does not match its bytes.');

      task = pdfjs.getDocument({ data: new Uint8Array(bytes), stopAtErrors: true, verbosity: 0,
        maxImageSize: 20_000_000, useSystemFonts: false });
      const doc = await task.promise;
      if (doc.numPages > limits.maxPages) throw new PdfOfficeError('LIMIT_EXCEEDED', 'PDF page count exceeds maxPages.');
      const source = { id: input.artifactRef.id, sha256: digest };

      if (operation === 'inspect') {
        const { info: rawInfo } = await doc.getMetadata();
        const info = rawInfo as Record<string, unknown>;
        const result = { format: 'pdf', mediaType: 'application/pdf', version: typeof info.PDFFormatVersion === 'string' ? info.PDFFormatVersion : null,
          pageCount: doc.numPages, encrypted: Boolean(info.EncryptFilterName), fingerprint: doc.fingerprints[0] ?? null,
          metadata: { title: typeof info.Title === 'string' ? info.Title : null, author: typeof info.Author === 'string' ? info.Author : null,
            creator: typeof info.Creator === 'string' ? info.Creator : null, producer: typeof info.Producer === 'string' ? info.Producer : null }, coverage: 'partial' };
        return { moduleId: PDF_OFFICE_MODULE_ID, requestId: input.requestId, operation, result, artifacts: [],
          warnings: [{ code: 'PDF_FEATURE_SCAN_PARTIAL', severity: 'info', message: 'This module reads PDF metadata and native text only; active-content scanning is not claimed.' }] as Warning[] };
      }

      if (operation === 'verify') {
        const plan = z.object({ expectedPageCount: z.number().int().positive().optional(), textIncludes: z.array(z.string().max(100_000)).optional() }).strict().safeParse(input.payload ?? {});
        if (!plan.success) throw new PdfOfficeError('INVALID_INPUT', plan.error.message);
        const checks: { id: string; status: 'pass' | 'fail' | 'skip'; severity: 'info' | 'warn' | 'error'; message: string }[] = [];
        checks.push({ id: 'structure.pdfjs', status: 'pass', severity: 'error', message: 'PDF reopened successfully using PDF.js.' });
        if (plan.data.expectedPageCount !== undefined) checks.push({ id: 'content.pageCount', status: plan.data.expectedPageCount === doc.numPages ? 'pass' : 'fail', severity: 'error', message: 'Page count compared with expectation.' });
        if (plan.data.textIncludes?.length) {
          const fullText = await extractText(doc, controller.signal, limits);
          plan.data.textIncludes.forEach((expected, index) => checks.push({ id: `content.text.${index}`, status: fullText.includes(expected) ? 'pass' : 'fail', severity: 'error', message: 'Expected text compared with the native text layer.' }));
        }
        const failed = checks.filter(check => check.status === 'fail').length;
        const skipped = checks.filter(check => check.status === 'skip').length;
        const report = { ok: failed === 0, partial: skipped > 0, checks,
          summary: { total: checks.length, passed: checks.length - failed - skipped, failed, skipped } };
        return { moduleId: PDF_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: report, verification: report, artifacts: [], warnings: [] as Warning[] };
      }

      const result = { source, engine: `pdfjs-${pdfjs.version}`, coordinates: 'pdf-user-space-points', pageCount: doc.numPages,
        pages: await extractPages(doc, controller.signal, limits), coverage: { text: 'native-text-only', readingOrder: 'not-inferred', tables: 'not-inferred', ocr: 'not-performed' } };
      if (Buffer.byteLength(JSON.stringify(result)) > limits.maxOutputJsonBytes)
        throw new PdfOfficeError('LIMIT_EXCEEDED', 'Parsed PDF exceeds maxOutputJsonBytes.');
      return { moduleId: PDF_OFFICE_MODULE_ID, requestId: input.requestId, operation, result, artifacts: [],
        warnings: [{ code: 'NATIVE_TEXT_ONLY', severity: 'info', message: 'PDF.js text items and page geometry are exposed; reading order, OCR, table semantics and visual quality are not inferred.' }] as Warning[] };
    } catch (error) {
      controller.signal.throwIfAborted();
      if (error instanceof PdfOfficeError) throw error;
      if (isPasswordError(error)) throw new PdfOfficeError('PASSWORD_REQUIRED', 'Password-protected PDF requires a dedicated credential workflow.');
      const message = error instanceof Error ? error.message : 'PDF.js could not process the PDF.';
      throw new PdfOfficeError('PARSE_FAILED', `PDF.js could not process the PDF: ${message.slice(0, 500)}`);
    } finally {
      clearTimeout(timer);
      await task?.destroy().catch(() => {});
    }
  };

  return { definition: PDF_OFFICE_DEFINITION, handlers: {
    inspect: (input: unknown) => invoke(input, 'inspect'),
    execute: (input: unknown) => invoke(input, 'execute'),
    verify: (input: unknown) => invoke(input, 'verify'),
  }, async dispose() { disposed = true; } };
}

async function extractPages(doc: PDFDocumentProxy, signal: AbortSignal, limits: PdfOfficeLimits) {
  const pages: { pageNumber: number; widthPt: number; heightPt: number; rotation: number; text: string; fragments: { text: string; transform: number[]; fontName: string; hasEOL: boolean }[] }[] = [];
  let textChars = 0, textItems = 0;
  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber++) {
    signal.throwIfAborted();
    const page = await doc.getPage(pageNumber), viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const fragments = content.items.flatMap(item => {
      if (!('str' in item) || !item.str) return [];
      const entry = { text: item.str, transform: [...item.transform], fontName: item.fontName, hasEOL: item.hasEOL };
      textChars += item.str.length; textItems++;
      if (textChars > limits.maxTextChars || textItems > limits.maxTextItems) throw new PdfOfficeError('LIMIT_EXCEEDED', 'PDF native text exceeds configured budgets.');
      return [entry];
    });
    pages.push({ pageNumber, widthPt: viewport.width, heightPt: viewport.height, rotation: page.rotate,
      text: fragments.map(fragment => fragment.text + (fragment.hasEOL ? '\n' : '')).join(''), fragments });
    page.cleanup();
  }
  return pages;
}
async function extractText(doc: PDFDocumentProxy, signal: AbortSignal, limits: PdfOfficeLimits) {
  return (await extractPages(doc, signal, limits)).map(page => page.text).join('\n');
}

export async function register(registry: { registerModule(module: ReturnType<typeof createPdfOfficeModule>): void | Promise<void> }, options: PdfOfficeOptions) {
  const module = createPdfOfficeModule(options); await registry.registerModule(module); return module;
}

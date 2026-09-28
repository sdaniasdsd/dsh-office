import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { ArtifactRef, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import { z } from 'zod';

export const PPTX_OFFICE_MODULE_ID = 'pptx-office' as const;
export const PPTX_OFFICE_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export const PPTX_OFFICE_DEFINITION = Object.freeze({
  id: PPTX_OFFICE_MODULE_ID, version: '0.1.0', profileGroup: 'PPTX', capabilities: PPTX_OFFICE_CAPABILITIES,
  summary: 'Inspect/extract PPTX text, replace addressed runs, apply guarded typography, and apply a cover-only artistic style without changing slide text.',
  dependencies: [{ name: 'python-pptx', kind: 'runtime' }, { name: 'office-core', kind: 'module' },
    { name: 'office-files', kind: 'module' }, { name: 'office-safety', kind: 'module' }],
  configSchema: { type: 'object', properties: { engine: { type: 'object', properties: { pythonPath: { type: 'string', default: 'python' } } } } },
});

const LIMITS = Object.freeze({ maxInputBytes: 64 * 1024 * 1024, maxOutputBytes: 64 * 1024 * 1024,
  maxOutputJsonBytes: 8 * 1024 * 1024, maxArchiveEntries: 4096, maxEntryUncompressedBytes: 64 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024, maxSlides: 500, maxShapes: 100_000, maxTextChars: 2_000_000, maxChanges: 5000 });
const artifactSchema = z.object({ id: z.string().min(1), uri: z.string().min(1), sha256: z.string().regex(/^[a-f0-9]{64}$/).optional(),
  sizeBytes: z.number().int().nonnegative().optional(), mediaType: z.string().optional(), label: z.string().optional() }).passthrough();
const changeSchema = z.object({ slideNumber: z.number().int().min(1), shapeId: z.number().int().min(0), paragraphIndex: z.number().int().min(0),
  runIndex: z.number().int().min(0), expectedText: z.string().min(1).max(100_000), replaceWith: z.string().max(100_000) }).strict();
const formatTextSchema = z.object({ action: z.literal('formatText'), changes: z.array(z.object({
  scope: z.literal('allSlides'), titleFontSize: z.number().min(8).max(96), bodyFontSize: z.number().min(8).max(72),
  accentColor: z.string().regex(/^#[0-9a-fA-F]{6}$/),
}).strict()).min(1).max(1) }).strict();
const artStyleSchema = z.object({ action: z.literal('applyArtStyle'), styleId: z.literal('indigo-paperlight-v1'),
  artWord: z.string().min(1).max(24).regex(/^[\p{L}\p{N} ._-]+$/u).optional() }).strict();
const requestSchema = z.object({ requestId: z.string().min(1), operation: z.enum(PPTX_OFFICE_CAPABILITIES), artifactRef: artifactSchema,
  policy: z.object({ id: z.string().min(1) }).passthrough().optional(), payload: z.record(z.unknown()).optional() }).passthrough();
const checkSchema = z.object({ id: z.string().min(1), status: z.enum(['pass', 'fail', 'skip']), severity: z.enum(['info', 'warn', 'error']), message: z.string() }).strict();
const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

export interface PptxArtifactStore {
  read(ref: ArtifactRef, maxBytes?: number, signal?: AbortSignal): Promise<Uint8Array>;
  write(input: { bytes: Uint8Array; suggestedName: string; requestId?: string; source?: ArtifactRef; sources?: ArtifactRef[] }): Promise<ArtifactRef>;
}
export interface PptxOfficeOptions { artifactStore: PptxArtifactStore; pythonPath?: string;
  safetyGuard?: { assertAllowed(input: { artifactRef: ArtifactRef; operation: string; policy: SafetyPolicy }): Promise<void> } }

class PptxOfficeError extends Error {
  constructor(readonly code: string, message: string) { super(message); this.name = 'PptxOfficeError'; }
}
function assertRef(ref: ArtifactRef) {
  if (extname(ref.label ?? ref.uri.split(/[?#]/, 1)[0] ?? '').toLowerCase() !== '.pptx')
    throw new PptxOfficeError('FORMAT_MISMATCH', 'pptx-office accepts macro-free .pptx files only.');
}
function runPython(command: string, script: string, operation: string, inputPath: string, outputPath: string, request: unknown, timeoutMs: number) {
  return new Promise<{ code: number; stdout: string; stderr: string }>((resolve, reject) => {
    const child = spawn(command, [script, operation, inputPath, outputPath], {
      windowsHide: true, shell: false, stdio: ['pipe', 'pipe', 'pipe'],
      // The bridge emits JSON with ensure_ascii=False; make the subprocess protocol
      // independent of the Windows user's legacy ANSI code page (commonly GBK).
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    });
    let stdout = '', stderr = '', size = 0, settled = false;
    const finish = (error?: unknown, result?: { code: number; stdout: string; stderr: string }) => {
      if (settled) return; settled = true; clearTimeout(timer);
      if (error) reject(error); else resolve(result!);
    };
    const timer = setTimeout(() => { child.kill(); finish(new PptxOfficeError('ENGINE_TIMEOUT', 'python-pptx exceeded the module time limit.')); }, timeoutMs);
    const read = (which: 'stdout' | 'stderr', chunk: Buffer) => {
      size += chunk.length;
      if (size > LIMITS.maxOutputJsonBytes + 1024 * 1024) { child.kill(); finish(new PptxOfficeError('LIMIT_EXCEEDED', 'PPTX engine output exceeded its byte limit.')); return; }
      if (which === 'stdout') stdout += chunk.toString('utf8'); else stderr += chunk.toString('utf8');
    };
    child.stdout.on('data', (chunk: Buffer) => read('stdout', chunk)); child.stderr.on('data', (chunk: Buffer) => read('stderr', chunk));
    child.once('error', () => finish(new PptxOfficeError('ENGINE_UNAVAILABLE', 'Could not start the configured Python runtime.')));
    child.once('close', code => finish(undefined, { code: code ?? -1, stdout, stderr }));
    child.stdin.end(JSON.stringify({ limits: LIMITS, ...(request as object) }));
  });
}

export function createPptxOfficeModule(options: PptxOfficeOptions) {
  if (!options?.artifactStore) throw new PptxOfficeError('INVALID_INPUT', 'artifactStore is required.');
  let disposed = false;
  const invoke = async (inputRaw: unknown, operation: 'inspect' | 'execute' | 'verify') => {
    const parsed = requestSchema.safeParse(inputRaw);
    if (!parsed.success || parsed.data.operation !== operation) throw new PptxOfficeError('INVALID_INPUT', parsed.success ? 'Handler operation mismatch.' : parsed.error.message);
    const input = parsed.data;
    if (disposed) throw new PptxOfficeError('MODULE_DISPOSED', 'PPTX module has been disposed.');
    assertRef(input.artifactRef);
    if (input.policy) {
      if (!options.safetyGuard) throw new PptxOfficeError('SAFETY_POLICY_DENIED', 'A safety policy was supplied without a configured guard.');
      await options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, operation, policy: input.policy as SafetyPolicy });
    }
    let bytes: Uint8Array;
    try { bytes = await options.artifactStore.read(input.artifactRef, LIMITS.maxInputBytes); }
    catch { throw new PptxOfficeError('ARTIFACT_NOT_FOUND', 'PPTX artifact could not be read from managed storage.'); }
    if (bytes.length > LIMITS.maxInputBytes) throw new PptxOfficeError('LIMIT_EXCEEDED', 'PPTX input exceeds byte budget.');
    const digest = sha256(bytes);
    if ((input.artifactRef.sha256 && input.artifactRef.sha256 !== digest) ||
        (input.artifactRef.sizeBytes !== undefined && input.artifactRef.sizeBytes !== bytes.length))
      throw new PptxOfficeError('ARTIFACT_INTEGRITY_FAILED', 'PPTX artifact metadata does not match its bytes.');
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new PptxOfficeError('FORMAT_MISMATCH', 'PPTX must be a ZIP-based OOXML package.');
    const temp = await mkdtemp(join(tmpdir(), 'dsh-pptx-'));
    try {
      const inputPath = join(temp, 'input.pptx'), outputPath = join(temp, 'output.pptx');
      await writeFile(inputPath, bytes);
      const payload = input.payload ?? {};
      const action = operation === 'inspect' ? 'inspect' : operation === 'verify' ? 'verify' : (payload as Record<string, unknown>).action;
      if (operation === 'execute' && !['extract', 'replaceText', 'formatText', 'applyArtStyle'].includes(String(action))) throw new PptxOfficeError('INVALID_INPUT', 'execute payload.action must be extract, replaceText, formatText, or applyArtStyle.');
      if (action === 'replaceText') {
        const check = z.object({ action: z.literal('replaceText'), changes: z.array(changeSchema).min(1).max(LIMITS.maxChanges) }).strict().safeParse(payload);
        if (!check.success) throw new PptxOfficeError('INVALID_INPUT', check.error.message);
      }
      if (action === 'formatText') {
        const check = formatTextSchema.safeParse(payload);
        if (!check.success) throw new PptxOfficeError('INVALID_INPUT', check.error.message);
      }
      if (action === 'applyArtStyle') {
        const check = artStyleSchema.safeParse(payload);
        if (!check.success) throw new PptxOfficeError('INVALID_INPUT', check.error.message);
      }
      const request = operation === 'execute' && ['replaceText', 'formatText'].includes(String(action))
        ? { changes: (payload as { changes: unknown }).changes } : payload;
      const result = await runPython(options.pythonPath ?? 'python', fileURLToPath(new URL('./engine/pptx_bridge.py', import.meta.url)),
        String(action), inputPath, outputPath, request, 60_000);
      if (result.code !== 0) {
        if (/No module named ['"]?pptx/.test(result.stderr)) throw new PptxOfficeError('ENGINE_UNAVAILABLE', 'python-pptx is not installed in the configured Python runtime.');
        if (/exceeds the Profile|count is empty|precondition failed|not supported|not a PowerPoint|macro-free/i.test(result.stderr))
          throw new PptxOfficeError('UNSUPPORTED_OPERATION', result.stderr.trim().slice(-1000));
        throw new PptxOfficeError('ENGINE_FAILED', `python-pptx failed: ${result.stderr.trim().slice(-1000)}`);
      }
      let data: Record<string, unknown>;
      try { data = JSON.parse(result.stdout); } catch { throw new PptxOfficeError('ENGINE_PROTOCOL_ERROR', 'python-pptx returned invalid JSON.'); }
      if (Buffer.byteLength(result.stdout) > LIMITS.maxOutputJsonBytes) throw new PptxOfficeError('LIMIT_EXCEEDED', 'PPTX result exceeds JSON byte budget.');
      if (operation === 'inspect') return { moduleId: PPTX_OFFICE_MODULE_ID, requestId: input.requestId, operation,
        result: data, artifacts: [], warnings: [] as Warning[] };
      if (operation === 'verify') {
        const checks = z.array(checkSchema).safeParse(data.checks);
        if (!checks.success || checks.data.length === 0) throw new PptxOfficeError('ENGINE_PROTOCOL_ERROR', 'Verifier returned no valid checks.');
        const failed = checks.data.filter(check => check.status === 'fail').length;
        const skipped = checks.data.filter(check => check.status === 'skip').length;
        const report = { ok: failed === 0, partial: skipped > 0, checks: checks.data,
          summary: { total: checks.data.length, passed: checks.data.length - failed - skipped, failed, skipped } };
        return { moduleId: PPTX_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: report, verification: report, artifacts: [], warnings: [] as Warning[] };
      }
      if (action === 'replaceText' || action === 'formatText' || action === 'applyArtStyle') {
        const outputBytes = new Uint8Array(await readFile(outputPath));
        if (!outputBytes.length || outputBytes.length > LIMITS.maxOutputBytes) throw new PptxOfficeError('LIMIT_EXCEEDED', 'Generated PPTX is empty or exceeds the output byte budget.');
        const source = { ...input.artifactRef, sha256: digest, sizeBytes: bytes.length };
        const artifactRef = await options.artifactStore.write({ bytes: outputBytes, requestId: input.requestId, sources: [source], suggestedName: 'edited.pptx' });
        return { moduleId: PPTX_OFFICE_MODULE_ID, requestId: input.requestId, operation,
          result: { artifactRef, sources: [{ id: input.artifactRef.id, sha256: digest }],
            ...(action === 'replaceText' ? { changedRuns: data.changedRuns } : action === 'applyArtStyle' ? {
              styleId: data.styleId, styledSlides: data.styledSlides, artWord: data.artWord,
            } : {
              formattedRuns: data.formattedRuns, formattedSlides: data.formattedSlides,
              verifiedFormattingRuns: data.verifiedFormattingRuns, contrastAdjustedSlides: data.contrastAdjustedSlides,
              titleColorPreservedSlides: data.titleColorPreservedSlides,
            }),
            slideCount: data.slideCount, visualReview: 'pending' }, artifacts: [artifactRef],
          warnings: [{ code: 'VISUAL_REVIEW_PENDING', severity: 'info', message: 'The deck was reopened and text preservation was checked; render and visually review every slide before delivery. Artistic styles are limited to the cover slide and do not alter themes, masters, charts, tables, or slide text.' }] };
      }
      return { moduleId: PPTX_OFFICE_MODULE_ID, requestId: input.requestId, operation,
        result: { source: { id: input.artifactRef.id, sha256: digest }, ...data, visualReview: 'pending' }, artifacts: [],
        warnings: [{ code: 'PPTX_OBSERVATIONS', severity: 'info', message: 'Shape IDs are source-revision-scoped; this extraction is not a visual layout review or an edit authorization.' }] };
    } finally { await rm(temp, { recursive: true, force: true }); }
  };
  return { definition: PPTX_OFFICE_DEFINITION, handlers: {
    inspect: (input: unknown) => invoke(input, 'inspect'), execute: (input: unknown) => invoke(input, 'execute'), verify: (input: unknown) => invoke(input, 'verify'),
  }, async dispose() { disposed = true; } };
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPptxOfficeModule>): void | Promise<void> }, options: PptxOfficeOptions) {
  const module = createPptxOfficeModule(options); await registry.registerModule(module); return module;
}

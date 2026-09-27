import { createHash } from 'node:crypto';
import type { ArtifactRef } from 'office-core';
import type { ModuleCreateOptions, DocxCreateModule, ModuleConfig, ModuleOutput, EngineResult } from './contract';
import { resolveConfig } from './config';
import { DOCX_CREATE_SCHEMAS, parseInput, executeSchema, inspectSchema, verifySchema } from './schema';
import { Docx4jCreateEngine } from './engine/adapter';
import { checkPackage, DOCX_MIME } from './engine/package-check';
import { failure, normalizeError, DocxCreateError } from './errors';
import { toFormatProfile, toCreateResult, pendingWarnings } from './mapper';
import { verifyPackage } from './verifier';
import { recordTelemetry } from './telemetry';
import { artifactRefSchema } from './domain/docx-create';

export const DOCX_CREATE_DEFINITION = {
  id: 'docx-create', version: '0.1.0', profileGroup: 'DOCX', capabilities: ['inspect', 'execute', 'verify'] as const,
  configSchema: DOCX_CREATE_SCHEMAS.config,
  dependencies: ['office-core', 'office-safety', 'office-files', 'office-test-kit'].map(name => ({ name, kind: 'module' })),
};
const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
async function engineDeadline<T>(work: () => Promise<T>, timeoutMs: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const start = Date.now();
  try {
    const value = await Promise.race([
      work(), new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new DocxCreateError('ENGINE_TIMEOUT', 'Engine exceeded its time budget.')), timeoutMs); }),
    ]);
    if (Date.now() - start >= timeoutMs) failure('ENGINE_TIMEOUT', 'Engine exceeded its time budget.');
    return value;
  } finally { if (timer) clearTimeout(timer); }
}
export function createDocxCreateModule(options: ModuleCreateOptions): DocxCreateModule {
  if (!options?.artifactStore?.read || !options.artifactStore.write) failure('INVALID_INPUT', 'An office-files artifact store must be injected.');
  const base = resolveConfig(options.config); const engine = options.engine ?? new Docx4jCreateEngine();
  let disposed = false;
  const run = async <T>(operation: string, task: () => Promise<T>): Promise<T> => {
    const start = Date.now(); let errorCode: string | undefined;
    try { if (disposed) failure('MODULE_DISPOSED', 'Module has been disposed.'); return await task(); }
    catch (error) { const e = normalizeError(error); errorCode = e.code; throw e; }
    finally { recordTelemetry(options.telemetry, { moduleId: 'docx-create', operation, elapsedMs: Date.now() - start, ok: !errorCode, ...(errorCode ? { errorCode } : {}) }); }
  };
  const authorize = async (input: Parameters<NonNullable<ModuleCreateOptions['safetyGuard']>['assertAllowed']>[0]) => {
    try { await options.safetyGuard?.assertAllowed(input); } catch { failure('SAFETY_POLICY_DENIED', 'Profile safety guard rejected this operation.'); }
  };
  const read = async (ref: ArtifactRef, config: ModuleConfig, request: { operation: string; policy?: Parameters<typeof authorize>[0]['policy'] }) => {
    await authorize({ artifactRef: ref, operation: request.operation, ...(request.policy ? { policy: request.policy } : {}) });
    let bytes: Uint8Array;
    try { bytes = await options.artifactStore.read(ref, config.limits.maxInputBytes); }
    catch (error) { if (error instanceof DocxCreateError) throw error; return failure('ARTIFACT_STORE_UNAVAILABLE', 'Artifact could not be read.'); }
    if (!(bytes instanceof Uint8Array) || bytes.length > config.limits.maxInputBytes) failure('LIMIT_EXCEEDED', 'Input artifact exceeds byte limit.');
    if ((ref.sha256 && ref.sha256 !== sha256(bytes)) || (ref.sizeBytes !== undefined && ref.sizeBytes !== bytes.length)) failure('ARTIFACT_INTEGRITY_FAILED', 'Artifact metadata does not match its bytes.');
    return bytes;
  };
  const output = <T>(requestId: string, operation: ModuleOutput<T>['operation'], result: T): ModuleOutput<T> =>
    ({ moduleId: 'docx-create', requestId, operation, result, artifacts: [], warnings: [] });
  return {
    definition: DOCX_CREATE_DEFINITION,
    handlers: {
      inspect: raw => run('inspect', async () => {
        const input = parseInput(inspectSchema, raw); const config = resolveConfig(input.options, base, true);
        const bytes = await read(input.artifactRef, config, input);
        return output(input.requestId, 'inspect', toFormatProfile(checkPackage(bytes, config, input.policy)));
      }),
      verify: raw => run('verify', async () => {
        const input = parseInput(verifySchema, raw); const config = resolveConfig(input.options, base, true);
        const bytes = await read(input.artifactRef, config, input);
        const report = verifyPackage(bytes, config, input.policy);
        return { ...output(input.requestId, 'verify', report), verification: report };
      }),
      execute: raw => run('execute', async () => {
        const input = parseInput(executeSchema, raw); const config = resolveConfig(input.options, base, true);
        await authorize({ operation: 'execute', ...(input.policy ? { policy: input.policy } : {}) });
        if (JSON.stringify(input.plan).length > config.limits.maxTextChars) failure('LIMIT_EXCEEDED', 'Creation specification exceeds configured character budget.');
        const images: Record<string, Uint8Array> = Object.create(null);
        const sources: ArtifactRef[] = []; let template: Uint8Array | undefined; let totalInputBytes = 0;
        if (input.plan.kind === 'fillTemplate') {
          if (!config.featureFlags.allowTemplateFill) failure('SAFETY_POLICY_DENIED', 'Template filling is disabled.');
          template = await read(input.plan.templateRef, config, input); sources.push(input.plan.templateRef);
          checkPackage(template, config, input.policy);
        } else {
          if (input.plan.document.blocks.length > Math.min(config.limits.maxBlocks, input.policy?.maxBlocks ?? Infinity)) failure('LIMIT_EXCEEDED', 'Block limit exceeded.');
          for (const b of input.plan.document.blocks) {
            if (b.kind === 'table' && b.rows.length * b.rows[0]!.length > Math.min(config.limits.maxTableCells, input.policy?.maxTableCells ?? Infinity)) failure('LIMIT_EXCEEDED', 'Table cell limit exceeded.');
            if (b.kind === 'image') {
              if (!config.featureFlags.allowImages) failure('SAFETY_POLICY_DENIED', 'Images are disabled.');
              const bytes = await read(b.artifactRef, config, input); totalInputBytes += bytes.length;
              if (totalInputBytes > config.limits.maxInputBytes) failure('LIMIT_EXCEEDED', 'Total image bytes exceed input limit.');
              images[b.id] = bytes; sources.push(b.artifactRef);
            }
          }
        }
        const result: EngineResult = await engineDeadline(() => engine.execute({ plan: input.plan, config, images, ...(template ? { template } : {}) }), config.timeoutMs);
        if (!(result?.bytes instanceof Uint8Array) || !Array.isArray(result.filledKeys) || result.filledKeys.some(k => typeof k !== 'string')) failure('ENGINE_FAILED', 'Engine returned an invalid result.');
        if (result.bytes.length > config.limits.maxOutputBytes) failure('LIMIT_EXCEEDED', 'Output artifact exceeds byte limit.');
        const facts = checkPackage(result.bytes, config, input.policy);
        const report = verifyPackage(result.bytes, config, input.policy);
        if (!report.ok) failure('VERIFICATION_FAILED', 'Output failed local verification.');
        if (disposed) failure('MODULE_DISPOSED', 'Module was disposed before artifact commit.');
        let ref: ArtifactRef;
        try {
          ref = await options.artifactStore.write({ requestId: input.requestId, sources, bytes: result.bytes, suggestedName: 'document.docx' });
        } catch { return failure('ARTIFACT_STORE_UNAVAILABLE', 'Artifact store could not commit output.'); }
        if (!artifactRefSchema.safeParse(ref).success || sources.some(s => s.id === ref.id || s.uri === ref.uri)) failure('ARTIFACT_STORE_UNAVAILABLE', 'Artifact store must return a new valid immutable artifact reference.');
        const digest = sha256(result.bytes);
        if ((ref.sha256 && ref.sha256 !== digest) || (ref.sizeBytes !== undefined && ref.sizeBytes !== result.bytes.length)) failure('ARTIFACT_INTEGRITY_FAILED', 'Artifact store returned inconsistent output metadata.');
        // Render's public adapter uses label before URI when selecting the DOCX file type.
        ref = { ...ref, label: ref.label?.toLowerCase().endsWith('.docx') ? ref.label : 'document.docx', sha256: digest, sizeBytes: result.bytes.length, mediaType: DOCX_MIME };
        const summary = toCreateResult(ref, engine.name, input.plan.kind, result.filledKeys, facts);
        return { ...output(input.requestId, 'execute', result.design ? { ...summary, design: result.design } : summary),
          artifacts: [ref], warnings: pendingWarnings(facts.fields), verification: report };
      }),
    },
    async dispose() { if (!disposed) { disposed = true; await engine.dispose(); } },
  };
}
export * from './contract';
export * from './domain/docx-create';
export { DOCX_CREATE_SCHEMAS } from './schema';
export { DocxCreateError, ERROR_CODES } from './errors';
export { PRESETS } from './engine/presets';

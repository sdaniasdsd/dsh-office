import type { z } from 'zod';
import type { ArtifactRef, EngineContext, ModuleDefinition, ModuleOptions, ModuleOutput, ModuleRequest,
  Operation, PdfFormatProfile, VerificationReport } from './contract';
import { requestSchema, engineEnvelopeSchema, profileSchema, reportSchema } from './schema';
import { PdfModuleError } from './errors';
import { resolveConfig, withCallOverrides } from './config';
import { assertJson, sha256, makeReport } from './validation';
import { assertSameSource, validateDualIR } from './ir';

export function createPdfModule<T>(definition: ModuleDefinition, resultSchema: z.ZodType<T>, options: ModuleOptions = {}) {
  const config = resolveConfig(options.config);
  let disposed = false;
  const activeControllers = new Set<AbortController>();

  async function run(operation: Operation, input: ModuleRequest): Promise<ModuleOutput<unknown>> {
    const started = Date.now();
    let errorCode: string | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const controller = new AbortController();
    try {
      if (disposed) throw new PdfModuleError('MODULE_DISPOSED', 'Module has been disposed.');
      assertJson(input);
      const parsed = requestSchema.safeParse(input);
      if (!parsed.success) throw new PdfModuleError('INVALID_INPUT', parsed.error.message);
      const request = parsed.data;
      if (request.operation !== operation) throw new PdfModuleError('INVALID_INPUT', 'Operation does not match handler.');
      if (!request.artifactRef && !(definition.id === 'pdf-create' && operation === 'execute'))
        throw new PdfModuleError('INVALID_INPUT', 'This operation requires artifactRef.');
      const effective = withCallOverrides(config, request.options);
      const engine = options.engine;
      if (!engine) throw new PdfModuleError('ENGINE_UNAVAILABLE', `${definition.id}: connect an artifact reader for the default backend, or inject an engine.`);
      if (request.policy && !options.safetyGuard)
        throw new PdfModuleError('SAFETY_POLICY_DENIED', 'Policy supplied without an injected safety guard.');
      activeControllers.add(controller);
      const abortPromise = new Promise<never>((_, reject) => {
        controller.signal.addEventListener('abort', () => reject(controller.signal.reason), { once: true });
      });
      timer = setTimeout(() => controller.abort(new PdfModuleError('ENGINE_TIMEOUT', 'Module operation timed out.')), effective.timeoutMs);
      const read = async (ref: ArtifactRef, budget = effective.limits.maxInputBytes): Promise<Uint8Array> => {
        controller.signal.throwIfAborted();
        if (!options.files) throw new PdfModuleError('ARTIFACT_NOT_FOUND', 'Artifact reader is not connected.');
        if (ref.sizeBytes !== undefined && ref.sizeBytes > budget)
          throw new PdfModuleError('LIMIT_EXCEEDED', 'Declared artifact exceeds read budget.');
        let bytes: Uint8Array;
        try { bytes = await options.files.read(ref, budget, controller.signal); }
        catch (error) {
          if (error instanceof PdfModuleError) throw error;
          throw new PdfModuleError('ARTIFACT_NOT_FOUND', 'Artifact reader failed.');
        }
        controller.signal.throwIfAborted();
        if (!(bytes instanceof Uint8Array)) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Reader must return Uint8Array.');
        if (bytes.byteLength > budget) throw new PdfModuleError('LIMIT_EXCEEDED', 'Artifact exceeds read budget.');
        if ((ref.sizeBytes !== undefined && ref.sizeBytes !== bytes.byteLength) || (ref.sha256 !== undefined && ref.sha256 !== sha256(bytes)))
          throw new PdfModuleError('ARTIFACT_MISMATCH', 'Artifact size or hash does not match its bytes.');
        return bytes;
      };
      const work = async (): Promise<ModuleOutput<unknown>> => {
        const context: EngineContext = { moduleId: definition.id, request, config: effective, signal: controller.signal };
        if (request.artifactRef) {
          const bytes = await read(request.artifactRef);
          context.source = { ref: request.artifactRef, bytes, identity: { id: request.artifactRef.id, sha256: sha256(bytes) } };
        }
        if (request.policy) {
          try { await options.safetyGuard!.assertAllowed({ ...context, policy: request.policy }); }
          catch { throw new PdfModuleError('SAFETY_POLICY_DENIED', 'The injected guard rejected this operation.'); }
        }
        controller.signal.throwIfAborted();
        const raw = await engine.invoke(context);
        controller.signal.throwIfAborted();
        assertJson(raw, 'ENGINE_PROTOCOL_ERROR');
        if (Buffer.byteLength(JSON.stringify(raw)) > effective.limits.maxOutputJsonBytes)
          throw new PdfModuleError('LIMIT_EXCEEDED', 'Engine JSON exceeds output budget.');
        const envelope = engineEnvelopeSchema.safeParse(raw);
        if (!envelope.success) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', envelope.error.message);
        const schema = operation === 'inspect' ? profileSchema : operation === 'verify' ? reportSchema : resultSchema;
        const result = schema.safeParse(envelope.data.result);
        if (!result.success) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', result.error.message);
        const data = result.data as Record<string, unknown>;
        if ('pageCount' in data && typeof data.pageCount === 'number' && data.pageCount > effective.limits.maxPages)
          throw new PdfModuleError('LIMIT_EXCEEDED', 'Page count exceeds Profile budget.');
        if (operation === 'verify') {
          const report = result.data as VerificationReport;
          if (!report.checks.length || new Set(report.checks.map(c => c.id)).size !== report.checks.length)
            throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Verification requires nonempty, uniquely named checks.');
          const calculated = makeReport(report.checks);
          if (report.ok !== calculated.ok || report.partial !== calculated.partial ||
              (['total', 'passed', 'failed', 'skipped'] as const).some(k => report.summary[k] !== calculated.summary[k]))
            throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Verification summary contradicts its checks.');
        }
        if (operation === 'execute' && 'ir' in data) {
          const ir = validateDualIR(data.ir);
          if (ir.pages.length > effective.limits.maxPages || ir.physical.length + ir.semantic.length + (ir.tables ?? []).reduce((sum, t) => sum + t.cells.length, 0) > effective.limits.maxNodes)
            throw new PdfModuleError('LIMIT_EXCEEDED', 'IR exceeds page or node budget.');
          if (context.source) assertSameSource(ir.source, context.source.identity);
        }
        if (operation === 'execute' && 'source' in data && context.source)
          assertSameSource(data.source as { id: string; sha256: string }, context.source.identity);
        if (operation === 'execute' && Array.isArray(data.pages)) {
          const pages = data.pages as { pageNumber: number }[];
          if (pages.length > effective.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'Page array exceeds budget.');
          if (pages.some((page, i) => page.pageNumber !== i + 1) || ('pageCount' in data && pages.length !== data.pageCount))
            throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Expected complete, ordered, one-based pages.');
        }
        const artifacts = envelope.data.artifacts;
        if (artifacts.length > effective.limits.maxArtifacts) throw new PdfModuleError('LIMIT_EXCEEDED', 'Too many artifacts.');
        if (new Set(artifacts.map(a => a.id)).size !== artifacts.length || new Set(artifacts.map(a => a.uri)).size !== artifacts.length)
          throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Duplicate artifact identity.');
        if (operation === 'execute') {
          const referenced: ArtifactRef[] = [];
          if ('artifactRef' in data) referenced.push(data.artifactRef as ArtifactRef);
          if ('manifestRef' in data) referenced.push(data.manifestRef as ArtifactRef);
          if (definition.id === 'pdf-render') {
            for (const page of data.pages as { image: ArtifactRef }[]) referenced.push(page.image);
          }
          for (const artifact of referenced) {
            const listed = artifacts.find(a => a.id === artifact.id);
            if (!listed || listed.uri !== artifact.uri || listed.sha256 !== artifact.sha256 || listed.sizeBytes !== artifact.sizeBytes)
              throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Result artifact must be declared with matching URI, hash and size.');
          }
          if (definition.id === 'pdf-edit' && context.source &&
              !(data.sources as { id: string; sha256: string }[]).some(s => s.id === context.source!.identity.id && s.sha256 === context.source!.identity.sha256))
            throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Edit result does not reference the input revision.');
          if (definition.id === 'pdf-artifact' && context.source)
            assertSameSource(data.document as { id: string; sha256: string }, context.source.identity);
        }
        let totalOutputBytes = 0;
        for (const artifact of artifacts) {
          if (!artifact.sha256 || artifact.sizeBytes === undefined)
            throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Outputs require size and SHA-256.');
          if (context.source && (artifact.id === context.source.ref.id || artifact.uri === context.source.ref.uri))
            throw new PdfModuleError('ARTIFACT_MISMATCH', 'Output must use a new identity and URI.');
          totalOutputBytes += artifact.sizeBytes;
          if (totalOutputBytes > effective.limits.maxTotalOutputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Output artifacts exceed total byte budget.');
          await read(artifact, effective.limits.maxOutputBytes);
        }
        const output: ModuleOutput<unknown> = { moduleId: definition.id, operation, requestId: request.requestId,
          result: result.data, artifacts, warnings: envelope.data.warnings };
        if (operation === 'verify') output.verification = result.data as VerificationReport;
        return output;
      };
      return await Promise.race([work(), abortPromise]);
    } catch (error) {
      const normalized = error instanceof PdfModuleError ? error : new PdfModuleError('ENGINE_FAILED', 'Injected engine or guard failed.', { cause: error });
      errorCode = normalized.code;
      throw normalized;
    } finally {
      if (timer) clearTimeout(timer);
      activeControllers.delete(controller);
      try {
        const recorded = options.telemetry?.record({ moduleId: definition.id, operation, ok: errorCode === undefined,
          elapsedMs: Date.now() - started, ...(errorCode ? { errorCode } : {}) });
        void Promise.resolve(recorded).catch(() => {});
      } catch { /* Telemetry never changes the document operation. */ }
    }
  }
  return {
    definition,
    handlers: {
      inspect: (input: ModuleRequest) => run('inspect', input) as Promise<ModuleOutput<PdfFormatProfile>>,
      execute: (input: ModuleRequest) => run('execute', input) as Promise<ModuleOutput<T>>,
      verify: (input: ModuleRequest) => run('verify', input) as Promise<ModuleOutput<VerificationReport>>,
    },
    async dispose() {
      if (disposed) return;
      disposed = true;
      for (const controller of activeControllers) controller.abort(new PdfModuleError('MODULE_DISPOSED', 'Module disposed during operation.'));
      await options.engine?.dispose?.();
    },
  };
}

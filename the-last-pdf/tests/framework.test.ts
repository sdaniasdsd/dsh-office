import { describe, it, expect, vi } from 'vitest';
import { readFile } from 'node:fs/promises';
import { createPdfInspectModule } from '../pdf-inspect/src/index';
import { createPdfParseModule } from '../pdf-parse/src/index';
import { createPdfCreateModule } from '../pdf-create/src/index';
import { createPdfRenderModule } from '../pdf-render/src/index';
import { createPdfEditModule } from '../pdf-edit/src/index';
import { createPdfArtifactModule } from '../pdf-artifact/src/index';
import { createPdfEasyParseModule } from '../pdf-easy-parse/src/index';
import { createPdfComplexParseModule } from '../pdf-complex-parse/src/index';
import { assertJson, makeReport, verifyPolicy, resolveConfig, withCallOverrides } from '@dsh-office-profile/pdf-contracts';
import { ref, bytes, reader, engine, profile, source, irFixture } from './support';

const request = { artifactRef: ref, requestId: 'test', operation: 'execute' as const };
describe('migrated module framework', () => {
  it.each([createPdfInspectModule, createPdfEasyParseModule, createPdfParseModule,
    createPdfComplexParseModule, createPdfCreateModule, createPdfEditModule, createPdfRenderModule, createPdfArtifactModule])(
    'requires an explicit engine for every module', async factory => {
      const module = factory();
      const manifest = JSON.parse(await readFile(new URL(`../${module.definition.id}/module.json`, import.meta.url), 'utf8'));
      expect(module.definition).toEqual(manifest);
      await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'ENGINE_UNAVAILABLE' });
    });
  it('reads bytes and returns a lossless envelope', async () => {
    const module = createPdfInspectModule({ files: reader, engine: engine(profile) });
    const output = await module.handlers.execute(request);
    expect(output.result).toEqual(profile);
    expect(JSON.parse(JSON.stringify(output))).toEqual(output);
  });
  it('does not confuse a supplied MIME or filename with verified bytes', async () => {
    const invoke = vi.fn();
    const module = createPdfInspectModule({ files: reader, engine: { name: 'spy', invoke } });
    await expect(module.handlers.execute({ ...request, artifactRef: { ...ref, sha256: '0'.repeat(64) } }))
      .rejects.toMatchObject({ code: 'ARTIFACT_MISMATCH' });
    expect(invoke).not.toHaveBeenCalled();
  });
  it('rejects oversize input even if the reader ignores the budget', async () => {
    const module = createPdfInspectModule({ files: reader, engine: engine(profile), config: { limits: { maxInputBytes: 1 } } });
    await expect(module.handlers.execute({ ...request, artifactRef: { id: ref.id, uri: ref.uri } }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  });
  it('rejects engine protocol violations and page limit overruns', async () => {
    await expect(createPdfInspectModule({ files: reader, engine: engine({ ...profile, encrypted: 'no' }) }).handlers.execute(request))
      .rejects.toMatchObject({ code: 'ENGINE_PROTOCOL_ERROR' });
    await expect(createPdfInspectModule({ files: reader, engine: engine({ ...profile, pageCount: 201 }) }).handlers.execute(request))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  });
  it('rejects a mismatched handler operation', async () => {
    await expect(createPdfInspectModule().handlers.inspect(request)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
  it('preserves Profile ceilings on request overrides', () => {
    const config = resolveConfig({ limits: { maxPages: 10 }, timeoutMs: 100 });
    expect(withCallOverrides(config, { limits: { maxPages: 2 } }).limits.maxPages).toBe(2);
    expect(() => withCallOverrides(config, { limits: { maxPages: 11 } })).toThrow('relax');
    expect(() => withCallOverrides(config, { timeoutMs: 101 })).toThrow('relax');
  });
  it('requires a safety guard when a policy is supplied', async () => {
    const module = createPdfInspectModule({ engine: engine(profile), files: reader });
    await expect(module.handlers.execute({ ...request, policy: { id: 'strict', allowJavaScript: false } }))
      .rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
  });
  it('keeps unknown features distinct from absent and undeclared', () => {
    expect(verifyPolicy(profile, { id: 'strict', allowEncrypted: false }).ok).toBe(false);
    expect(verifyPolicy(profile).summary.skipped).toBe(5);
    expect(verifyPolicy(profile).partial).toBe(true);
  });
  it('rejects contradictory verification summaries', async () => {
    const report = makeReport([{ id: 'parse', status: 'fail', severity: 'error', message: 'Malformed' }]);
    report.ok = true;
    await expect(createPdfInspectModule({ files: reader, engine: engine(report) }).handlers.verify({ ...request, operation: 'verify' }))
      .rejects.toMatchObject({ code: 'ENGINE_PROTOCOL_ERROR' });
  });
  it('does not interpret skipped visual review as complete', async () => {
    const report = makeReport([{ id: 'visual', status: 'skip', severity: 'info', message: 'Not reviewed' }]);
    const output = await createPdfInspectModule({ files: reader, engine: engine(report) }).handlers.verify({ ...request, operation: 'verify' });
    expect(output.verification).toEqual(report);
    expect(output.result.partial).toBe(true);
  });
  it('times out a stalled engine and aborts its context', async () => {
    let signal: AbortSignal | undefined;
    const module = createPdfInspectModule({ files: reader, config: { timeoutMs: 15 }, engine: {
      name: 'stalled', async invoke(context) { signal = context.signal; return new Promise(() => {}); },
    } });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'ENGINE_TIMEOUT' });
    expect(signal?.aborted).toBe(true);
  });
  it('disposal interrupts ongoing operations and is idempotent', async () => {
    let entered!: () => void;
    const started = new Promise<void>(resolve => { entered = resolve; });
    const dispose = vi.fn(async () => {});
    const module = createPdfInspectModule({ files: reader, engine: {
      name: 'stalled', async invoke() { entered(); return new Promise(() => {}); }, dispose,
    } });
    const pending = module.handlers.execute(request);
    const rejected = expect(pending).rejects.toMatchObject({ code: 'MODULE_DISPOSED' });
    await started;
    await module.dispose();
    await rejected;
    await module.dispose();
    expect(dispose).toHaveBeenCalledTimes(1);
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'MODULE_DISPOSED' });
  });
  it('telemetry failures do not fail a document operation', async () => {
    const module = createPdfInspectModule({ files: reader, engine: engine(profile),
      telemetry: { record: async () => { throw new Error('offline'); } } });
    expect((await module.handlers.execute(request)).result).toEqual(profile);
  });
  it('binds IR to actual source bytes, including refs without a hash', async () => {
    const ir = irFixture();
    ir.source.sha256 = 'f'.repeat(64);
    const module = createPdfParseModule({ files: reader, engine: engine({ ir }) });
    await expect(module.handlers.execute({ ...request, artifactRef: { id: ref.id, uri: ref.uri } }))
      .rejects.toMatchObject({ code: 'SOURCE_MAP_MISMATCH' });
  });
  it('allows source-free creation to reach the selected engine', async () => {
    const invoke = vi.fn(async () => { throw new Error('synthetic sentinel'); });
    await expect(createPdfCreateModule({ engine: { name: 'synthetic', invoke } }).handlers.execute({ requestId: 'new', operation: 'execute', payload: { title: 'new' } }))
      .rejects.toMatchObject({ code: 'ENGINE_FAILED' });
    expect(invoke.mock.calls).toHaveLength(1);
  });
  it('rejects reordered easy-parse pages', async () => {
    const module = createPdfEasyParseModule({ files: reader, engine: engine({ source, pages: [{ pageNumber: 2, text: 'x', coverage: 'partial' }] }) });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'ENGINE_PROTOCOL_ERROR' });
  });
  it('requires output references to appear in the artifact list', async () => {
    const artifactRef = { ...ref, id: 'output', uri: 'memory:output', mediaType: 'application/pdf' };
    const module = createPdfEditModule({ files: reader, engine: engine({ artifactRef, sources: [source], visualReview: 'pending' }) });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'ENGINE_PROTOCOL_ERROR' });
  });
  it('refuses source identity reuse for produced artifacts', async () => {
    const artifactRef = { ...ref, mediaType: 'application/pdf' };
    const module = createPdfEditModule({ files: reader, engine: { name: 'synthetic', async invoke() {
      return { result: { artifactRef, sources: [source], visualReview: 'pending' }, artifacts: [artifactRef], warnings: [] };
    } } });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'ARTIFACT_MISMATCH' });
  });
  it('reopens output references and detects corrupted bytes', async () => {
    const artifactRef = { ...ref, id: 'output', uri: 'memory:output', sha256: '0'.repeat(64), mediaType: 'application/pdf' };
    const module = createPdfCreateModule({ files: reader, engine: { name: 'synthetic', async invoke() {
      return { result: { artifactRef, sources: [], visualReview: 'pending' }, artifacts: [artifactRef], warnings: [] };
    } } });
    await expect(module.handlers.execute({ requestId: 'create', operation: 'execute' })).rejects.toMatchObject({ code: 'ARTIFACT_MISMATCH' });
  });
  it('accepts an independently stored output with matching digest', async () => {
    const artifactRef = { ...ref, id: 'output', uri: 'memory:output', mediaType: 'application/pdf' };
    const module = createPdfEditModule({ files: reader, engine: { name: 'synthetic', async invoke() {
      return { result: { artifactRef, sources: [source], visualReview: 'pending' }, artifacts: [artifactRef], warnings: [] };
    } } });
    expect((await module.handlers.execute(request)).artifacts).toEqual([artifactRef]);
  });
  it('rejects editing evidence for a different source revision', async () => {
    const artifactRef = { ...ref, id: 'output', uri: 'memory:output', mediaType: 'application/pdf' };
    const module = createPdfEditModule({ files: reader, engine: { name: 'synthetic', async invoke() {
      return { result: { artifactRef, sources: [{ ...source, sha256: 'f'.repeat(64) }], visualReview: 'pending' }, artifacts: [artifactRef], warnings: [] };
    } } });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'SOURCE_MAP_MISMATCH' });
  });
  it('maps guard rejection to a policy error and skips the engine', async () => {
    const invoke = vi.fn();
    const module = createPdfInspectModule({ files: reader, engine: { name: 'spy', invoke },
      safetyGuard: { async assertAllowed() { throw new Error('denied'); } } });
    await expect(module.handlers.execute({ ...request, policy: { id: 'deny' } })).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
    expect(invoke).not.toHaveBeenCalled();
  });
  it('enforces JSON output byte budget', async () => {
    const module = createPdfInspectModule({ files: reader, engine: engine(profile), config: { limits: { maxOutputJsonBytes: 1 } } });
    await expect(module.handlers.execute(request)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  });
  it.each([NaN, Infinity, undefined, new Date(), bytes, { x: undefined }, [undefined], new Array(2)])('rejects lossy JSON %#', value => {
    expect(() => assertJson(value)).toThrow();
  });
  it('rejects cyclic JSON but accepts reused ordinary values', () => {
    const a: Record<string, unknown> = {};
    a.self = a;
    expect(() => assertJson(a)).toThrow();
    const b = { x: 1 };
    expect(() => assertJson([b, b])).not.toThrow();
  });
});

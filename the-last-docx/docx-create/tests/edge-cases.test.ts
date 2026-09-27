import { describe, expect, it, vi } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import { generated, memoryStore, simple, rewrite, part, asBytes } from './helpers';
import type { ExecuteInput } from '../src/contract';

describe('rejection and limits', () => {
  it.each([null, { ...simple, requestId: '' }, { ...simple, callback() {} }, { ...simple, options: { timeoutMs: NaN } }, { ...simple, timestamp: new Date() }])('rejects invalid or lossy JSON input', async input => {
    const m = memoryStore();
    await expect(createDocxCreateModule({ artifactStore: m.store }).handlers.execute(input as ExecuteInput)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    expect(m.files.size).toBe(0);
  });
  it('rejects duplicate IDs, bad XML text and nonrectangular tables', async () => {
    const m = createDocxCreateModule({ artifactStore: memoryStore().store });
    for (const blocks of [
      [{ kind: 'paragraph', id: 'same', runs: [{ text: 'a' }] }, { kind: 'pageBreak', id: 'same' }],
      [{ kind: 'paragraph', id: 'bad', runs: [{ text: '\u0001' }] }],
      [{ kind: 'table', id: 'table', rows: [['a'], ['b', 'c']] }],
    ]) await expect(m.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks } } } as ExecuteInput)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
  it('prevents per-call policy escalation and honors the safety guard', async () => {
    const memory = memoryStore();
    const m = createDocxCreateModule({ artifactStore: memory.store, config: { featureFlags: { allowTemplateFill: false } } });
    await expect(m.handlers.execute({ ...simple, options: { featureFlags: { allowTemplateFill: true } } })).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
    const denied = createDocxCreateModule({ artifactStore: memory.store, safetyGuard: { async assertAllowed() { throw new Error('no'); } } });
    await expect(denied.handlers.execute(simple)).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
    expect(memory.files.size).toBe(0);
  });
  it('rejects corrupt ZIPs, external relationships and broken targets', async () => {
    const { bytes, put, module } = await generated();
    const variants = [
      [asBytes('not a docx'), 'FORMAT_MISMATCH'],
      [rewrite(bytes, e => { e['word/_rels/document.xml.rels'] = asBytes(part(bytes, 'word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="external" Type="hyperlink" TargetMode="External" Target="https://example.com"/></Relationships>')); }), 'SAFETY_POLICY_DENIED'],
      [rewrite(bytes, e => { delete e['word/styles.xml']; }), 'VERIFICATION_FAILED'],
      [rewrite(bytes, e => { e['word/document.xml'] = asBytes('<!DOCTYPE x [<!ENTITY xx "x">]><x/>'); }), 'SAFETY_POLICY_DENIED'],
    ] as const;
    for (const [data, code] of variants) await expect(module.handlers.inspect({ operation: 'inspect', requestId: 'bad', artifactRef: put('bad', data) })).rejects.toMatchObject({ code });
  });
  it('enforces input/output/expanded limits and does not commit invalid bytes', async () => {
    const { bytes } = await generated(); const memory = memoryStore(); const ref = memory.put('doc', bytes);
    for (const limits of [{ maxInputBytes: 1 }, { maxExpandedBytes: 100 }, { maxEntries: 1 }]) {
      const module = createDocxCreateModule({ artifactStore: memory.store, config: { limits } });
      await expect(module.handlers.inspect({ operation: 'inspect', requestId: 'limit', artifactRef: ref })).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    }
    const write = vi.spyOn(memory.store, 'write');
    const module = createDocxCreateModule({ artifactStore: memory.store, config: { limits: { maxOutputBytes: 100 } } });
    await expect(module.handlers.execute(simple)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    expect(write).not.toHaveBeenCalled();
  });
  it('reports failed/partial verification, not success, for a corrupt document', async () => {
    const memory = memoryStore(); const ref = memory.put('bad', asBytes('broken'));
    const out = await createDocxCreateModule({ artifactStore: memory.store }).handlers.verify({ operation: 'verify', requestId: 'bad', artifactRef: ref });
    expect(out.result).toMatchObject({ ok: false, partial: true, summary: { failed: 1 } });
  });
  it('rejects stale input hashes', async () => {
    const { module, out } = await generated();
    await expect(module.handlers.inspect({ operation: 'inspect', requestId: 'stale', artifactRef: { ...out.result.artifactRef, sha256: '0'.repeat(64) } })).rejects.toMatchObject({ code: 'ARTIFACT_INTEGRITY_FAILED' });
  });
  it('does not commit late engine output after timeout', async () => {
    const { bytes } = await generated(); const memory = memoryStore(); const write = vi.spyOn(memory.store, 'write');
    const module = createDocxCreateModule({ artifactStore: memory.store, config: { timeoutMs: 5 }, engine: {
      name: 'slow', async execute() { await new Promise(resolve => setTimeout(resolve, 30)); return { bytes, filledKeys: [] }; }, async dispose() {},
    } });
    await expect(module.handlers.execute(simple)).rejects.toMatchObject({ code: 'ENGINE_TIMEOUT' });
    await new Promise(resolve => setTimeout(resolve, 50)); expect(write).not.toHaveBeenCalled();
  });
  it('telemetry cannot turn a completed write into a failure', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store, telemetry: { record() { throw new Error('telemetry down'); } } });
    expect((await module.handlers.execute(simple)).artifacts).toHaveLength(1);
  });
});

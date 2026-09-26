import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { ArtifactRef } from 'office-core';
import { createDocxParseModule } from '@dsh-office-profile/docx-parse';
import type { ArtifactStore, EngineRenderResult, RenderEngine } from '../src/contract';
import { createDocxRenderModule } from '../src/index';

const sourceRef: ArtifactRef = {
  id: 'source-1', uri: 'file:///fixtures/report.docx', label: 'report.docx',
  mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  sizeBytes: 4,
};
const sourceBytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04]);
const parseFixture = fileURLToPath(new URL('../../docx-parse/fixtures/_generated/clean.docx', import.meta.url));

class MemoryStore implements ArtifactStore {
  readonly stored: Array<{ ref: ArtifactRef; bytes: Uint8Array }> = [];
  constructor(private readonly bytes = sourceBytes) {}
  async read(): Promise<Uint8Array> { return this.bytes; }
  async write(input: { source: ArtifactRef; requestId: string; bytes: Uint8Array; suggestedName: string }): Promise<ArtifactRef> {
    const ref = { id: `generated-${this.stored.length + 1}`, uri: `memory://${input.requestId}/${input.suggestedName}`, label: input.suggestedName };
    this.stored.push({ ref, bytes: input.bytes });
    return ref;
  }
}

class FakeEngine implements RenderEngine {
  readonly name = 'fake-renderer';
  renderCalls = 0;
  async inspect() { return { extension: '.docx', mediaType: sourceRef.mediaType ?? null }; }
  async render(): Promise<EngineRenderResult> {
    this.renderCalls += 1;
    return {
      pdf: new Uint8Array([37, 80, 68, 70]),
      pages: [
        { pageNumber: 1, image: new Uint8Array([137, 80, 78, 71]), thumbnail: new Uint8Array([137, 80, 78, 71]) },
        { pageNumber: 2, image: new Uint8Array([137, 80, 78, 71]), thumbnail: new Uint8Array([137, 80, 78, 71]) },
      ],
    };
  }
  async dispose(): Promise<void> {}
}

describe('docx-render module contract', () => {
  it('inspects a DOCX using the injected engine and returns Profile-compatible metadata', async () => {
    const store = new MemoryStore();
    const module = createDocxRenderModule({ artifactStore: store, engine: new FakeEngine() });
    const result = await module.handlers.inspect({ artifactRef: sourceRef, requestId: 'inspect-1', operation: 'inspect' });
    expect(result.result).toMatchObject({ format: 'docx', extension: 'docx', container: 'zip' });
    expect(JSON.parse(JSON.stringify(result))).toEqual(result);
  });

  it('stores a PDF, each page image, and thumbnails as ArtifactStore artifacts', async () => {
    const store = new MemoryStore();
    const module = createDocxRenderModule({ artifactStore: store, engine: new FakeEngine() });
    const output = await module.handlers.execute({ artifactRef: sourceRef, requestId: 'render-1', operation: 'execute' });
    expect(output.result.pageCount).toBe(2);
    expect(output.result.pdf?.label).toBe('report.preview.pdf');
    expect(output.result.pages.map((page) => page.image.label)).toEqual(['report.page-1.png', 'report.page-2.png']);
    expect(output.result.pages.every((page) => page.thumbnail)).toBe(true);
    expect(output.artifacts).toHaveLength(5);
    expect(output.result.visualReview).toBe('pending');
  });

  it('does not claim visual verification until the agent explicitly supplies its review', async () => {
    const engine = new FakeEngine();
    const module = createDocxRenderModule({ artifactStore: new MemoryStore(), engine });
    const rendered = await module.handlers.execute({ artifactRef: sourceRef, requestId: 'render-2', operation: 'execute' });
    const pending = await module.handlers.verify({ artifactRef: sourceRef, requestId: 'verify-pending', operation: 'verify', renderResult: rendered.result });
    expect(pending.verification).toMatchObject({ ok: true, partial: true, summary: { skipped: 1 } });
    const reviewed = await module.handlers.verify({
      artifactRef: sourceRef, requestId: 'verify-reviewed', operation: 'verify', renderResult: rendered.result,
      visualFindings: [{ id: 'f1', kind: 'truncation', severity: 'error', page: 2, message: 'Paragraph is clipped.' }],
    });
    expect(reviewed.verification).toMatchObject({ ok: false, partial: false, summary: { failed: 1 } });
    expect(reviewed.warnings[0]).toMatchObject({ code: 'VISUAL_FINDINGS_PRESENT' });
    expect(engine.renderCalls).toBe(1);
  });

  it('rejects findings referring to pages outside the render result', async () => {
    const module = createDocxRenderModule({ artifactStore: new MemoryStore(), engine: new FakeEngine() });
    const rendered = await module.handlers.execute({ artifactRef: sourceRef, requestId: 'render-3', operation: 'execute' });
    await expect(module.handlers.verify({
      artifactRef: sourceRef, requestId: 'verify-bad-page', operation: 'verify', renderResult: rendered.result,
      visualFindings: [{ id: 'f1', kind: 'overlap', severity: 'warn', page: 3, message: 'Outside range.' }],
    })).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('rejects a render result from a different artifact revision', async () => {
    const module = createDocxRenderModule({ artifactStore: new MemoryStore(), engine: new FakeEngine() });
    const rendered = await module.handlers.execute({ artifactRef: sourceRef, requestId: 'render-stale', operation: 'execute' });
    await expect(module.handlers.verify({
      artifactRef: { ...sourceRef, id: 'new-revision' }, requestId: 'verify-stale', operation: 'verify', renderResult: rendered.result,
      visualFindings: [],
    })).rejects.toMatchObject({ code: 'STALE_RENDER_RESULT' });
  });

  it('validates visual finding references against the public docx-parse source map', async () => {
    const bytes = new Uint8Array(await readFile(parseFixture));
    const artifactRef = { ...sourceRef, uri: parseFixture, label: 'clean.docx', sizeBytes: bytes.byteLength };
    const parser = createDocxParseModule();
    const parsed = await parser.handlers.execute({ artifactRef, requestId: 'parse-for-render', operation: 'execute' });
    const currentArtifact = parsed.result.artifact;
    const sourceSemanticId = Object.keys(parsed.result.ir.content.sourceMap.bySemanticId)[0];
    if (!sourceSemanticId) throw new Error('The regression fixture should contain a semantic source node.');
    const renderer = createDocxRenderModule({ artifactStore: new MemoryStore(bytes), engine: new FakeEngine() });
    const rendered = await renderer.handlers.execute({ artifactRef: currentArtifact, requestId: 'render-parser-fixture', operation: 'execute' });
    const reviewed = await renderer.handlers.verify({
      artifactRef: currentArtifact, requestId: 'verify-source-map', operation: 'verify', renderResult: rendered.result,
      parseResult: parsed.result,
      visualFindings: [{ id: 'finding-source-map', kind: 'overlap', severity: 'warn', page: 1, message: 'Review this source node.', sourceSemanticIds: [sourceSemanticId] }],
    });
    expect(reviewed.warnings[0]?.details).toMatchObject({ sourceSemanticIds: [sourceSemanticId] });
    await expect(renderer.handlers.verify({
      artifactRef: currentArtifact, requestId: 'verify-missing-source-map', operation: 'verify', renderResult: rendered.result,
      visualFindings: [{ id: 'finding-source-map', kind: 'overlap', severity: 'warn', page: 1, message: 'Review this source node.', sourceSemanticIds: [sourceSemanticId] }],
    })).rejects.toMatchObject({ code: 'SOURCE_MAP_REQUIRED' });
    await parser.dispose();
  });

  it('uses the injected office-safety guard when a policy is supplied', async () => {
    const module = createDocxRenderModule({
      artifactStore: new MemoryStore(), engine: new FakeEngine(),
      safetyGuard: { async assertAllowed() { throw new Error('blocked by policy'); } },
    });
    await expect(module.handlers.execute({
      artifactRef: sourceRef, requestId: 'denied', operation: 'execute', policy: { id: 'strict' },
    })).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED', message: 'blocked by policy' });
  });

  it('reports a missing local renderer as an unavailable dependency, not a visual pass', async () => {
    const bytes = new Uint8Array(await readFile(parseFixture));
    const artifactRef = { ...sourceRef, uri: parseFixture, label: 'clean.docx', sizeBytes: bytes.byteLength };
    const module = createDocxRenderModule({
      artifactStore: new MemoryStore(bytes),
      config: { engine: { sofficePath: '__docx_render_missing_soffice__' } },
    });
    await expect(module.handlers.execute({ artifactRef, requestId: 'renderer-missing', operation: 'execute' }))
      .rejects.toMatchObject({ code: 'ENGINE_UNAVAILABLE' });
  });
});

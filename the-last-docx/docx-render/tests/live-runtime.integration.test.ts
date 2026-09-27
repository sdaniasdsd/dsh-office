import { readdir, readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import type { ArtifactRef } from 'office-core';
import { describe, expect, it } from 'vitest';
import type { ArtifactStore } from '../src/contract';
import { createDocxRenderModule } from '../src/index';

const runtimeRoot = process.env['DSH_LIVE_RUNTIME_ROOT'];

class RecordingStore implements ArtifactStore {
  readonly writes: Array<{ ref: ArtifactRef; bytes: Uint8Array }> = [];
  constructor(private readonly source: Uint8Array) {}
  async read(): Promise<Uint8Array> { return this.source; }
  async write(input: { source: ArtifactRef; requestId: string; bytes: Uint8Array; suggestedName: string }): Promise<ArtifactRef> {
    const ref: ArtifactRef = {
      id: `generated-${this.writes.length + 1}`,
      uri: `memory://${input.requestId}/${input.suggestedName}`,
      label: input.suggestedName,
      sizeBytes: input.bytes.byteLength,
    };
    this.writes.push({ ref, bytes: input.bytes });
    return ref;
  }
}

describe('live DSH runtime integration', () => {
  it.runIf(runtimeRoot !== undefined)('injects a DSH runtime root into the real handler and emits render artifacts', async () => {
    const generatedRoot = fileURLToPath(new URL('../../docx-create/fixtures/_generated/', import.meta.url));
    const fixtureName = (await readdir(generatedRoot, { withFileTypes: true }))
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name)
      .sort()[0];
    if (fixtureName === undefined) throw new Error(`No generated fixture directory under ${generatedRoot}`);
    const path = fileURLToPath(new URL(`../../docx-create/fixtures/_generated/${fixtureName}/technical-sample.docx`, import.meta.url));
    const bytes = new Uint8Array(await readFile(path));
    const source: ArtifactRef = {
      id: 'live-source', uri: path, label: 'technical-sample.docx',
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', sizeBytes: bytes.byteLength,
    };
    const store = new RecordingStore(bytes);
    const module = createDocxRenderModule({ artifactStore: store, config: { runtimeRoot } });
    try {
      const output = await module.handlers.execute({ operation: 'execute', artifactRef: source, requestId: 'live-dsh-runtime' });
      expect(output.result.engine).toBe('libreoffice-poppler');
      expect(output.result.pageCount).toBeGreaterThan(0);
      expect(output.result.pdf).toBeDefined();
      expect(store.writes.some((entry) => entry.ref.label?.endsWith('.pdf') === true)).toBe(true);
      expect(store.writes.some((entry) => entry.ref.label?.endsWith('.png') === true)).toBe(true);
    } finally {
      await module.dispose();
    }
  }, 120_000);
});


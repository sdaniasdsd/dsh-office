import { execFile as execFileCallback } from 'node:child_process';
import { existsSync } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import type { ArtifactRef } from 'office-core';
import { describe, expect, it } from 'vitest';
import type { ArtifactStore } from '../src/contract';
import { createDocxRenderModule } from '../src/index';

const runtimeRoot = process.env['DSH_LIVE_RUNTIME_ROOT'];
const execFile = promisify(execFileCallback);

async function generatedFixturePath(): Promise<string> {
  const createRoot = fileURLToPath(new URL('../../docx-create/', import.meta.url));
  const generatedRoot = fileURLToPath(new URL('../../docx-create/fixtures/_generated/', import.meta.url));
  let entries;
  try { entries = await readdir(generatedRoot, { withFileTypes: true }); }
  catch {
    // A clean checkout deliberately ignores generated documents.  Build the
    // test fixture here so enabling the live-runtime gate is self-contained.
    const tsx = [
      '../../../node_modules/tsx/dist/cli.mjs',
      '../../node_modules/tsx/dist/cli.mjs',
      '../../../dsh-office/node_modules/tsx/dist/cli.mjs',
    ].map((relative) => fileURLToPath(new URL(relative, import.meta.url))).find(existsSync);
    if (tsx === undefined) throw new Error('Cannot find the workspace tsx runner needed to build the live-render fixture');
    const builder = fileURLToPath(new URL('../../docx-create/fixtures/build.ts', import.meta.url));
    await execFile(process.execPath, [tsx, builder], { cwd: createRoot, windowsHide: true });
    entries = await readdir(generatedRoot, { withFileTypes: true });
  }
  const fixtureName = entries.filter((entry) => entry.isDirectory()).map((entry) => entry.name).sort().at(-1);
  if (fixtureName === undefined) throw new Error(`Fixture builder produced no directory under ${generatedRoot}`);
  return fileURLToPath(new URL(`../../docx-create/fixtures/_generated/${fixtureName}/technical-sample.docx`, import.meta.url));
}

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
    const path = await generatedFixturePath();
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

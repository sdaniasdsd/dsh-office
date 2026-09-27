import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import type { ArtifactRef } from 'office-core';
import { resolveConfig } from '../src/config';
import { LibreOfficePopplerEngine } from '../src/engine/adapter';

const fixturePath = process.env.DSH_RENDER_FIXTURE_PATH;
const sofficePath = process.env.DOCX_SOFFICE;
const pdftoppmPath = process.env.DOCX_PDFTOPPM;
const available = Boolean(fixturePath && sofficePath && pdftoppmPath);
const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

describe('LibreOffice profile reuse', () => {
  it.skipIf(!available)('serially reuses one isolated office profile without changing page or thumbnail pixels', async () => {
    const bytes = new Uint8Array(await readFile(fixturePath!));
    const artifact: ArtifactRef = { id: 'integration-fixture', uri: `file://${fixturePath}`, label: 'render-fixture.docx',
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', sizeBytes: bytes.length };
    const config = resolveConfig({ engine: { sofficePath: sofficePath!, pdftoppmPath: pdftoppmPath! } });
    const engine = new LibreOfficePopplerEngine(config.engine);
    try {
      const first = await engine.render(bytes, config, artifact);
      const second = await engine.render(bytes, config, artifact);
      expect(first.pages.map(page => ({ image: digest(page.image), thumbnail: page.thumbnail && digest(page.thumbnail) })))
        .toEqual(second.pages.map(page => ({ image: digest(page.image), thumbnail: page.thumbnail && digest(page.thumbnail) })));
      expect(second.pages.map(page => page.pageNumber)).toEqual(first.pages.map(page => page.pageNumber));
      expect(second.pages.length).toBeGreaterThan(0);
    } finally { await engine.dispose(); }
  }, 180_000);
});

import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createPptxOfficeModule } from '../src/index';

const python = 'python';
const available = spawnSync(python, ['-c', 'import pptx'], { windowsHide: true }).status === 0;
function store() {
  const objects = new Map<string, Uint8Array>();
  return {
    objects,
    async write(input: { bytes: Uint8Array; suggestedName: string }) {
      const hash = createHash('sha256').update(input.bytes).digest('hex');
      const ref = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/${input.suggestedName}`, sha256: hash, sizeBytes: input.bytes.length, label: input.suggestedName, mediaType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };
      objects.set(ref.uri, input.bytes); return ref;
    },
    async read(ref: { uri: string }) { const bytes = objects.get(ref.uri); if (!bytes) throw new Error('missing'); return bytes; },
  };
}

describe('pptx-office', () => {
  it.skipIf(!available)('inspects, extracts, and replaces one addressed run without rebuilding the deck', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'dsh-pptx-test-')), path = join(dir, 'source.pptx');
    try {
      const code = `from pptx import Presentation; from pptx.util import Inches; p=Presentation(); s=p.slides.add_slide(p.slide_layouts[6]); sh=s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)); sh.text="Before"; p.save(${JSON.stringify(path)})`;
      const generated = spawnSync(python, ['-c', code], { windowsHide: true, encoding: 'utf8' });
      expect(generated.status, generated.stderr).toBe(0);
      const bytes = new Uint8Array(await readFile(path)), hash = createHash('sha256').update(bytes).digest('hex');
      const source = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/source.pptx`, sha256: hash, sizeBytes: bytes.length, label: 'source.pptx' };
      const artifacts = store(); artifacts.objects.set(source.uri, bytes);
      const module = createPptxOfficeModule({ artifactStore: artifacts, pythonPath: python });
      const extracted = await module.handlers.execute({ requestId: 'extract-test', operation: 'execute', artifactRef: source, payload: { action: 'extract' } }) as unknown as { result: { slides: { shapes: { shapeId: number; paragraphs: { runs: { text: string }[] }[] }[] }[] } };
      const shape = extracted.result.slides[0]!.shapes[0]!;
      expect(shape.paragraphs[0]!.runs[0]!.text).toBe('Before');
      const edited = await module.handlers.execute({ requestId: 'replace-test', operation: 'execute', artifactRef: source, payload: { action: 'replaceText', changes: [{ slideNumber: 1, shapeId: shape.shapeId, paragraphIndex: 0, runIndex: 0, expectedText: 'Before', replaceWith: 'After' }] } }) as { result: { artifactRef: { uri: string } } };
      expect(artifacts.objects.has(edited.result.artifactRef.uri)).toBe(true);
      const verification = await module.handlers.verify({ requestId: 'verify-test', operation: 'verify', artifactRef: edited.result.artifactRef, payload: { expectedSlideCount: 1, textIncludes: ['After'] } }) as { result: { ok: boolean } };
      expect(verification.result.ok).toBe(true);
    } finally { await rm(dir, { recursive: true, force: true }); }
  });
});

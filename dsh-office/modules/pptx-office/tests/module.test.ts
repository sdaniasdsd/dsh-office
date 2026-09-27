import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createPptxOfficeModule } from '../src/index';

const python = process.env.DSH_DOCX_PYTHON ?? 'python';
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

type FormattedExtraction = {
  result: {
    slides: Array<{
      shapes: Array<{
        name: string;
        paragraphs: Array<{
          text: string;
          runs: Array<{ fontSizePt: number | null; fontColor: string | null }>;
        }>;
      }>;
    }>;
  };
};

describe('pptx-office', () => {
  it.skipIf(!available)('writes Unicode JSON as UTF-8 regardless of the inherited Windows Python encoding', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'dsh-pptx-unicode-')), path = join(dir, 'unicode.pptx');
    const text = 'Office\u00a0£ 中文';
    const oldIoEncoding = process.env.PYTHONIOENCODING;
    const oldUtf8Mode = process.env.PYTHONUTF8;
    try {
      // Force the pre-fix Windows failure mode even when the test host uses UTF-8.
      process.env.PYTHONIOENCODING = 'gbk';
      delete process.env.PYTHONUTF8;
      const code = `from pptx import Presentation; from pptx.util import Inches; p=Presentation(); s=p.slides.add_slide(p.slide_layouts[6]); sh=s.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1)); sh.text=${JSON.stringify(text)}; p.save(${JSON.stringify(path)})`;
      const generated = spawnSync(python, ['-c', code], { windowsHide: true, encoding: 'utf8' });
      expect(generated.status, generated.stderr).toBe(0);
      const bytes = new Uint8Array(await readFile(path)), hash = createHash('sha256').update(bytes).digest('hex');
      const source = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/unicode.pptx`, sha256: hash, sizeBytes: bytes.length, label: 'unicode.pptx' };
      const artifacts = store(); artifacts.objects.set(source.uri, bytes);
      const module = createPptxOfficeModule({ artifactStore: artifacts, pythonPath: python });
      const extracted = await module.handlers.execute({ requestId: 'unicode-extract-test', operation: 'execute', artifactRef: source, payload: { action: 'extract' } }) as unknown as { result: { slides: { shapes: { paragraphs: { text: string }[] }[] }[] } };
      expect(extracted.result.slides[0]!.shapes[0]!.paragraphs[0]!.text).toBe(text);
    } finally {
      if (oldIoEncoding === undefined) delete process.env.PYTHONIOENCODING; else process.env.PYTHONIOENCODING = oldIoEncoding;
      if (oldUtf8Mode === undefined) delete process.env.PYTHONUTF8; else process.env.PYTHONUTF8 = oldUtf8Mode;
      await rm(dir, { recursive: true, force: true });
    }
  });

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

  it.skipIf(!available)('formats slide title/body hierarchy without changing text', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'dsh-pptx-format-')), path = join(dir, 'source.pptx');
    try {
      const code = `from pptx import Presentation; from pptx.dml.color import RGBColor; from pptx.util import Inches, Pt; p=Presentation(); s=p.slides.add_slide(p.slide_layouts[1]); s.background.fill.solid(); s.background.fill.fore_color.rgb=RGBColor(255,255,255); s.shapes.title.text="Quarterly review"; s.shapes.title.text_frame.paragraphs[0].runs[0].font.size=Pt(40); s.placeholders[1].text="Revenue and risk"; s.placeholders[1].text_frame.paragraphs[0].runs[0].font.size=Pt(24); t=p.slides.add_slide(p.slide_layouts[1]); t.background.fill.solid(); t.background.fill.fore_color.rgb=RGBColor(100,20,100); t.shapes.title.text="Dark-background title"; t.shapes.title.height=Inches(0.3); t.shapes.title.text_frame.paragraphs[0].runs[0].font.size=Pt(40); t.placeholders[1].text="Body on colored slide"; t.placeholders[1].text_frame.paragraphs[0].runs[0].font.size=Pt(24); p.save(${JSON.stringify(path)})`;
      const generated = spawnSync(python, ['-c', code], { windowsHide: true, encoding: 'utf8' });
      expect(generated.status, generated.stderr).toBe(0);
      const bytes = new Uint8Array(await readFile(path)), hash = createHash('sha256').update(bytes).digest('hex');
      const source = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/source.pptx`, sha256: hash, sizeBytes: bytes.length, label: 'source.pptx' };
      const artifacts = store(); artifacts.objects.set(source.uri, bytes);
      const module = createPptxOfficeModule({ artifactStore: artifacts, pythonPath: python });
      const output = await module.handlers.execute({ requestId: 'format-test', operation: 'execute', artifactRef: source, payload: {
        action: 'formatText', changes: [{ scope: 'allSlides', titleFontSize: 30, bodyFontSize: 18, accentColor: '#244A67' }],
      } }) as { result: { artifactRef: { uri: string }; formattedRuns: number; verifiedFormattingRuns: number } };
      const edited = artifacts.objects.get(output.result.artifactRef.uri)!;
      const extracted = await module.handlers.execute({ requestId: 'extract-formatted', operation: 'execute', artifactRef: output.result.artifactRef, payload: { action: 'extract' } }) as unknown as FormattedExtraction;
      const texts = extracted.result.slides.flatMap(slide => slide.shapes.flatMap(shape => shape.paragraphs.map(paragraph => paragraph.text)));
      expect(texts).toContain('Quarterly review');
      expect(texts).toContain('Revenue and risk');
      expect(texts).toContain('Dark-background title');
      const firstSlide = extracted.result.slides[0]!;
      const secondSlide = extracted.result.slides[1]!;
      const title = firstSlide.shapes.find(shape => shape.name === 'Title 1');
      const body = firstSlide.shapes.find(shape => shape.name === 'Content Placeholder 2');
      const darkTitle = secondSlide.shapes.find(shape => shape.name === 'Title 1');
      if (!title || !body || !darkTitle) throw new Error('Expected title and body placeholders in both slides.');
      expect(title.paragraphs[0]!.runs[0]!).toMatchObject({ fontSizePt: 30, fontColor: '#244A67' });
      expect(body.paragraphs[0]!.runs[0]!.fontSizePt).toBe(18);
      expect(darkTitle.paragraphs[0]!.runs[0]!).toMatchObject({ fontSizePt: 40, fontColor: '#FFFFFF' });
      expect(output.result.verifiedFormattingRuns).toBeGreaterThan(0);
      expect(output.result.formattedRuns).toBeGreaterThan(0);
      expect(edited.length).toBeGreaterThan(0);
    } finally { await rm(dir, { recursive: true, force: true }); }
  });
});

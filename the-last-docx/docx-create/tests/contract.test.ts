import { describe, expect, it } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { createDocxCreateModule, DOCX_CREATE_SCHEMAS, DocxCreateError } from '../src/index';
import { generated, memoryStore, simple, part } from './helpers';
import { executeOutputSchema } from '../src/schema';
import manifest from '../module.json';

describe('public contract with real engine', () => {
  it('creates a DOCX with no input artifact, readable by the engine', async () => {
    const { out, bytes } = await generated();
    expect(out.result.artifactRef.sha256).toMatch(/^[a-f0-9]{64}$/);
    expect(out.result.paragraphIds).toHaveLength(2);
    expect(out.verification).toMatchObject({ ok: true, partial: true });
    expect(out.result).toMatchObject({ fields: 'pending-update', visualReview: 'pending' });
    expect(JSON.parse(JSON.stringify(out))).toEqual(out);
    expect(executeOutputSchema.safeParse(out).success).toBe(true);
    const pkg = await WordprocessingMLPackage.load(bytes);
    expect((await pkg.getBody()).paragraphs.map(p => p.text)).toContain('初始内容');
  });
  it('inspect and verify operate on the created artifact', async () => {
    const { module, out } = await generated(); const artifactRef = out.result.artifactRef;
    const inspected = await module.handlers.inspect({ requestId: 'inspect', operation: 'inspect', artifactRef });
    expect(inspected.result.format).toBe('docx');
    const verified = await module.handlers.verify({ requestId: 'verify', operation: 'verify', artifactRef });
    expect(verified.result.ok).toBe(true); expect(verified.result.partial).toBe(true);
    await module.dispose();
    await expect(module.handlers.execute(simple)).rejects.toMatchObject({ code: 'MODULE_DISPOSED' });
  });
  it('has machine-readable input/config schemas and serializable errors', () => {
    expect(DOCX_CREATE_SCHEMAS.execute).toHaveProperty('properties');
    expect(manifest.id).toBe('docx-create');
    expect(manifest.capabilities).toEqual(['inspect', 'execute', 'verify']);
    expect(Object.keys(manifest.configSchema.properties).sort()).toEqual(Object.keys((DOCX_CREATE_SCHEMAS.config as { properties: object }).properties).sort());
    expect(JSON.parse(JSON.stringify(new DocxCreateError('INVALID_INPUT', 'invalid')))).toEqual({ name: 'DocxCreateError', code: 'INVALID_INPUT', message: 'invalid' });
  });
  it('supports independent engine injection without changing handlers', async () => {
    const good = await generated(); const memory = memoryStore();
    const module = createDocxCreateModule({ artifactStore: memory.store, engine: {
      name: 'replacement-engine', async execute() { return { bytes: good.bytes, filledKeys: [] }; }, async dispose() {},
    } });
    const out = await module.handlers.execute(simple);
    expect(out.result.engine).toBe('replacement-engine');
  });
  it('creates headings, table, header/footer, TOC and explicit page break', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const result = await module.handlers.execute({ operation: 'execute', requestId: 'rich', plan: { kind: 'create', document: {
      preset: 'chinese-long', header: '项目简报', footer: '内部资料', blocks: [
        { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '测试文档' }] },
        { kind: 'toc', id: 'contents' }, { kind: 'pageBreak', id: 'break' },
        { kind: 'paragraph', id: 'chapter', style: 'Heading1', runs: [{ text: '第一章' }] },
        { kind: 'table', id: 'table', rows: [['项目', '说明'], ['测试', '内容']], columnWidthsMm: [40, 120] },
      ],
    } } });
    const bytes = memory.files.get(result.result.artifactRef.id)!;
    expect(part(bytes)).toContain('TOC'); expect(part(bytes)).toContain('w:tblHeader');
    expect(part(bytes)).toContain('w:type="page"');
    expect(part(bytes, 'word/header1.xml')).toContain('项目简报');
    expect(part(bytes, 'word/footer1.xml')).toContain('PAGE');
    expect(part(bytes, 'word/settings.xml')).toContain('w:updateFields');
  });
  it('embeds an image with a real internal relationship and preserves its bytes', async () => {
    const memory = memoryStore(); const module = createDocxCreateModule({ artifactStore: memory.store });
    const png = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=', 'base64'));
    const ref = memory.put('image', png);
    const out = await module.handlers.execute({ ...simple, plan: { kind: 'create', document: { blocks: [
      { kind: 'image', id: 'figure', artifactRef: ref, widthMm: 20, heightMm: 20, altText: '白色测试图片', caption: '图 1' },
    ] } } });
    const bytes = memory.files.get(out.result.artifactRef.id)!;
    expect(part(bytes)).toContain('r:embed'); expect(part(bytes)).toContain('白色测试图片');
    expect(out.verification?.ok).toBe(true);
    const { unzipSync } = await import('fflate');
    expect(unzipSync(bytes)['word/media/image1.png']).toEqual(png);
    await expect(module.handlers.execute({ ...simple, options: { featureFlags: { allowImages: false } }, plan: { kind: 'create', document: { blocks: [
      { kind: 'image', id: 'figure', artifactRef: ref, widthMm: 20, heightMm: 20, altText: '拒绝' },
    ] } } })).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
  });
});

import { describe, expect, it } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import type { DocumentSpec } from '../src/domain/docx-create';
import { memoryStore, part } from './helpers';

async function create(document: Partial<DocumentSpec>) {
  const memory = memoryStore();
  const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute({
    requestId: 'paragraph-format', operation: 'execute', plan: { kind: 'create', document: { blocks: document.blocks ?? [], ...document } },
  });
  const bytes = memory.files.get(out.result.artifactRef.id)!;
  return { out, body: part(bytes) };
}

/** The `w:pPr` of the paragraph whose text is exactly this. */
function paragraphPropertiesOf(body: string, text: string): string {
  const paragraph = body.split('</w:p>').find((piece) => piece.includes(`>${text}</w:t>`));
  if (paragraph === undefined) throw new Error(`no paragraph carrying ${text}`);
  return /<w:pPr>[\s\S]*?<\/w:pPr>/u.exec(paragraph)?.[0] ?? '';
}

describe('a paragraph may state its own first-line indent', () => {
  it('writes zero as a statement, so a flush-left line stops inheriting the style', async () => {
    const { body } = await create({ blocks: [
      { kind: 'paragraph', id: 'p1', runs: [{ text: '甲方（盖章）：' }], firstLineIndentPt: 0 },
    ] });
    // Both attributes, because Word lets the character count win over the length.
    expect(paragraphPropertiesOf(body, '甲方（盖章）：')).toContain('<w:ind w:firstLineChars="0" w:firstLine="0"/>');
  });

  it('converts points with the size the paragraph actually renders at', async () => {
    const { body } = await create({ blocks: [
      { kind: 'paragraph', id: 'p1', runs: [{ text: '两字缩进', size: 12 }], firstLineIndentPt: 24 },
      { kind: 'paragraph', id: 'p2', runs: [{ text: '五号缩进', size: 10.5 }], firstLineIndentPt: 21 },
    ] });
    expect(paragraphPropertiesOf(body, '两字缩进')).toContain('<w:ind w:firstLineChars="200" w:firstLine="480"/>');
    // 21 pt at 10.5 pt per character is still two characters: 200 hundredths.
    expect(paragraphPropertiesOf(body, '五号缩进')).toContain('<w:ind w:firstLineChars="200" w:firstLine="420"/>');
  });

  it('states nothing when the plan says nothing', async () => {
    const { body } = await create({ blocks: [
      { kind: 'paragraph', id: 'p1', runs: [{ text: '正文' }] },
    ] });
    expect(paragraphPropertiesOf(body, '正文')).not.toContain('<w:ind');
  });

  it('places w:ind where CT_PPr declares it: after spacing, before jc', async () => {
    const { body } = await create({ blocks: [
      { kind: 'paragraph', id: 'p1', runs: [{ text: '居中短行' }], alignment: 'center', firstLineIndentPt: 0 },
    ] });
    const ppr = paragraphPropertiesOf(body, '居中短行');
    expect(ppr.indexOf('<w:ind')).toBeLessThan(ppr.indexOf('<w:jc'));
  });

  it('rejects an indent no page could hold', async () => {
    await expect(create({ blocks: [
      { kind: 'paragraph', id: 'p1', runs: [{ text: '正文' }], firstLineIndentPt: 60 },
    ] })).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
});

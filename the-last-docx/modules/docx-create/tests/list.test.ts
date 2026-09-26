import { describe, expect, it } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import { DEFAULT_LIST_INDENT } from '../src/engine/numbering';
import type { DocumentSpec } from '../src/domain/docx-create';
import { memoryStore, part } from './helpers';

async function create(blocks: DocumentSpec['blocks'], document: Partial<DocumentSpec> = {}) {
  const memory = memoryStore();
  const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute({
    requestId: 'list-case', operation: 'execute', plan: { kind: 'create', document: { blocks, ...document } },
  });
  const bytes = memory.files.get(out.result.artifactRef.id)!;
  let numbering: string | undefined;
  try { numbering = part(bytes, 'word/numbering.xml'); } catch { numbering = undefined; }
  return { out, body: part(bytes), numbering };
}

const FOLLOW_UPS: DocumentSpec['blocks'] = [
  { kind: 'paragraph', id: 'h-next', style: 'Heading1', runs: [{ text: '后续跟进' }] },
  { kind: 'list', id: 'next', items: ['取得客户确认。', '完成 1 项修复中的延期事项。'] },
];

describe('list blocks become numbered content, not typed markers', () => {
  it('writes a numbering part and points each item at its definition', async () => {
    const { out, body, numbering } = await create(FOLLOW_UPS);
    expect(out.verification?.ok).toBe(true);
    expect(numbering).toBeDefined();
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>');
    expect(numbering).toContain('<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>');
    // Two items, two paragraphs, both naming the same definition.
    expect(body.match(/<w:numId w:val="1"\/>/gu)).toHaveLength(2);
    // The marker comes from the numbering definition; it is not document text.
    expect(body).not.toContain('\u2022');
    expect(body).not.toContain('•');
  });

  it('numbers an ordered list from the decimal definition', async () => {
    const { body, numbering } = await create([{ kind: 'list', id: 'steps', items: ['第一步', '第二步'], ordered: true }]);
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
    expect(numbering).toContain('<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>');
    expect(body.match(/<w:numId w:val="2"\/>/gu)).toHaveLength(2);
  });

  it('writes no numbering definitions for a document that has no lists', async () => {
    // The package ships an empty numbering part, so the question is not whether
    // the part exists but whether this module put anything in it.
    const { numbering } = await create([{ kind: 'paragraph', id: 'p', runs: [{ text: '没有列表' }] }]);
    expect(numbering ?? '').not.toContain('<w:abstractNum');
    expect(numbering ?? '').not.toContain('<w:num ');
  });

  it('keeps w:numPr after w:pStyle, the sequence ECMA-376 declares', async () => {
    const { body } = await create(FOLLOW_UPS);
    expect(/<w:pPr><w:pStyle w:val="Normal"\/><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="1"\/><\/w:numPr><\/w:pPr>/u.test(body)).toBe(true);
  });

  it('lets the register decide how far the markers hang', async () => {
    // A body with a two-character first-line indent needs its markers further
    // left, so item text lines up with the body rather than sitting proud of it.
    const academic = await create(FOLLOW_UPS, { scenario: 'academic-report', preset: 'chinese-long' });
    expect(academic.numbering).toContain('w:left="480" w:hanging="240"');

    const report = await create(FOLLOW_UPS, { scenario: 'internal-review' });
    expect(report.numbering).toContain(`w:left="${DEFAULT_LIST_INDENT.left}" w:hanging="${DEFAULT_LIST_INDENT.hanging}"`);
  });
});

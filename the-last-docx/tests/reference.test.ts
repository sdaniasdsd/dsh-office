import { describe, expect, it } from 'vitest';
import { normalizeText } from '@dsh-office-profile/docx-parse';
import type { CreatePlan } from '@dsh-office-profile/docx-create';
import { planFromContent, type SemanticContent } from '../src/reference';
import { context } from './support';

/** The plan's document, narrowed once so each test can read the blocks it cares about. */
interface PlanBlock { kind: string; id: string; style?: string; runs?: { text: string }[]; rows?: string[][]; columnWidthsMm?: number[]; items?: string[]; ordered?: boolean }
const docOf = (plan: CreatePlan) => (plan as unknown as { document: { blocks: PlanBlock[] } }).document;

/** A minimal dual IR, written by hand so each mapping rule can be pinned exactly. */
function content(blocks: unknown[], annotations: unknown[] = []): SemanticContent {
  return { semantic: { blocks: blocks as never, annotations: annotations as never } };
}

describe('text normalization keeps the ideographic space', () => {
  it('collapses layout whitespace but never eats U+3000', () => {
    // `\s` in JavaScript matches U+3000, so the earlier form turned the author's
    // `第一条　合同期限` into `第一条 合同期限` and the IR's text stopped being the
    // document's text. U+3000 is a character Word stores, not layout space.
    expect(normalizeText('第一条\u3000合同期限')).toBe('第一条\u3000合同期限');
    expect(normalizeText('\u3000首行缩进')).toBe('\u3000首行缩进');
    expect(normalizeText('a\t\tb\n c')).toBe('a b c');
    expect(normalizeText('  \u00a0 padded \n')).toBe('padded');
  });
});

describe('reference → creation plan', () => {
  it('carries body text, heading levels and tables', () => {
    const { plan, report } = planFromContent(content([
      { kind: 'heading', level: 1, text: '第一条　劳动合同期限' },
      { kind: 'paragraph', text: '1.1　本合同为固定期限劳动合同。' },
      { kind: 'table', gridColumns: 2, rows: [{ cells: [{ text: '甲方' }, { text: '乙方' }] }, { cells: [{ text: '甲值' }, { text: '乙值' }] }] },
    ]));
    const blocks = docOf(plan).blocks as { kind: string; style?: string; runs?: { text: string }[]; rows?: string[][] }[];
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'paragraph', 'table']);
    expect(blocks[0]!.style).toBe('Heading1');
    expect(blocks[0]!.runs?.[0]?.text).toBe('第一条　劳动合同期限');
    expect(blocks[1]!.style).toBe('Normal');
    expect(blocks[2]!.rows).toEqual([['甲方', '乙方'], ['甲值', '乙值']]);
    expect(report.carried).toMatchObject({ headings: 1, paragraphs: 1, tables: 1 });
  });

  it('re-lays a table out across the printable width instead of the reference’s geometry', () => {
    const { plan } = planFromContent(content([
      { kind: 'table', gridColumns: 4, rows: [{ cells: [{ text: 'a' }, { text: 'b' }, { text: 'c' }, { text: 'd' }] }] },
    ]), { tableWidthMm: 152 });
    const [table] = docOf(plan).blocks as { columnWidthsMm?: number[] }[];
    expect(table!.columnWidthsMm).toEqual([38, 38, 38, 38]);
    expect(table!.columnWidthsMm!.reduce((a, b) => a + b, 0)).toBe(152);
  });

  it('gathers a run of list-styled paragraphs into one real list block', () => {
    const { plan, report } = planFromContent(content([
      { kind: 'paragraph', styleId: 'ListBullet', text: '维护服务运行指标。' },
      { kind: 'paragraph', styleId: 'ListBullet', text: '参与变更评审。' },
      { kind: 'paragraph', text: '以上为岗位职责。' },
      { kind: 'paragraph', styleId: 'ListNumber', text: '第一步' },
    ]));
    const blocks = docOf(plan).blocks as { kind: string; items?: string[]; ordered?: boolean }[];
    expect(blocks.map((b) => b.kind)).toEqual(['list', 'paragraph', 'list']);
    expect(blocks[0]!.items).toEqual(['维护服务运行指标。', '参与变更评审。']);
    expect(blocks[0]!.ordered).toBeUndefined();
    expect(blocks[2]!.ordered).toBe(true);
    expect(report.carried).toMatchObject({ lists: 2, listItems: 3 });
  });

  it('maps only the styles a plan can name, and says which ones it could not', () => {
    const { plan, report } = planFromContent(content([
      { kind: 'paragraph', styleId: 'Title', text: '劳动合同书' },
      { kind: 'paragraph', styleId: 'Quote', text: '引用' },
      { kind: 'paragraph', styleId: 'Quote', text: '再引用' },
    ]));
    const blocks = docOf(plan).blocks as { style?: string }[];
    expect(blocks[0]!.style).toBe('Title');
    expect(blocks[1]!.style).toBe('Normal');
    expect(report.unmappedStyles).toEqual({ Quote: 2 });
    expect(report.notes.join(' ')).toMatch(/no creation-plan equivalent/u);
  });

  it('reports what a plan-built document cannot carry, instead of dropping it silently', () => {
    const { report } = planFromContent(content([
      { kind: 'paragraph', text: '正文' },
      { kind: 'image', altText: '图' },
    ], [{ kind: 'comment' }, { kind: 'footnote' }, { kind: 'revision' }]));
    expect(report.notCarriedByDesign).toEqual({ comments: 1, footnotes: 1, endnotes: 0, revisions: 1 });
    expect(report.skipped).toEqual([{ kind: 'image', reason: 'a creation plan has no block for this node kind' }]);
  });

  it('gives every block a plan-valid id', () => {
    const { plan } = planFromContent(content([
      { kind: 'paragraph', text: 'a' }, { kind: 'heading', level: 2, text: 'b' }, { kind: 'paragraph', text: 'c' },
    ]));
    const ids = (docOf(plan).blocks as { id: string }[]).map((b) => b.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const id of ids) expect(id).toMatch(/^[A-Za-z][A-Za-z0-9_.-]{0,79}$/u);
  });

  it('clamps a deeper heading level onto the three levels a plan defines', () => {
    const { plan } = planFromContent(content([{ kind: 'heading', level: 7, text: '深' }]));
    expect((docOf(plan).blocks as { style?: string }[])[0]!.style).toBe('Heading3');
  });
});

describe('reference → plan → document, against a real parse', () => {
  it('carries the text and structure of a document it just built', async () => {
    const c = await context();
    try {
      const source = await c.profile.call('docx-create', 'execute', {
        requestId: 'ref-source',
        operation: 'execute',
        plan: { kind: 'create', document: { preset: 'report', blocks: [
          { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '劳动合同书' }] },
          { kind: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: '第一条　合同期限' }] },
          { kind: 'paragraph', id: 'p1', runs: [{ text: '本合同为固定期限劳动合同。' }] },
          { kind: 'list', id: 'l1', items: ['取得客户确认。', '完成修复事项。'] },
          { kind: 'table', id: 't1', columnWidthsMm: [55, 97], rows: [['单位名称', '＿＿＿＿'], ['住所', '＿＿＿＿']] },
        ] } },
      }) as { result: { artifactRef: unknown } };

      const parsed = await c.profile.call('docx-parse', 'execute', {
        artifactRef: source.result.artifactRef, requestId: 'ref-parse',
      }) as { result: { ir: { content: SemanticContent } } };

      const { plan, report } = planFromContent(parsed.result.ir.content, { preset: 'chinese-contract', scenario: 'formal-record' });
      expect(report.carried.headings).toBe(1);
      expect(report.carried.tables).toBe(1);
      expect(report.carried.lists).toBe(1);
      expect(report.skipped).toEqual([]);

      const rebuilt = await c.profile.call('docx-create', 'execute', { requestId: 'ref-rebuild', operation: 'execute', plan }) as { result: { artifactRef: { sha256?: string } } };
      expect(rebuilt.result.artifactRef.sha256).toMatch(/^[a-f0-9]{64}$/u);

      // The rebuilt document parses to the same visible text and table shape.
      const again = await c.profile.call('docx-parse', 'execute', {
        artifactRef: rebuilt.result.artifactRef, requestId: 'ref-verify',
      }) as { result: { ir: { content: SemanticContent } } };
      const texts = (again.result.ir.content.semantic.blocks as { kind: string; text?: string; rows?: { cells: { text?: string }[] }[] }[])
        .filter((b) => b.kind !== 'table').map((b) => b.text);
      for (const expected of ['劳动合同书', '第一条　合同期限', '本合同为固定期限劳动合同。', '取得客户确认。']) {
        expect(texts).toContain(expected);
      }
      const table = (again.result.ir.content.semantic.blocks as { kind: string; rows?: { cells: { text?: string }[] }[] }[]).find((b) => b.kind === 'table');
      expect(table?.rows?.map((row) => row.cells.map((cell) => cell.text))).toEqual([['单位名称', '＿＿＿＿'], ['住所', '＿＿＿＿']]);
    } finally {
      await c.close();
    }
  });
});

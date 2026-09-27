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

describe('presentational observation is opt-in and is carried when asked for', () => {
  const FORMATTING = {
    paragraphs: [
      { blockId: 'b-title', alignment: 'center', runs: [{ text: '劳动合同书', bold: true, size: 22, eastAsia: '黑体' }] },
      { blockId: 'b-clause', alignment: null, runs: [{ text: '第一条　合同期限', bold: true, size: 12, eastAsia: '黑体' }] },
      { blockId: 'b-no', alignment: 'right', runs: [{ text: '编号：＿＿＿' }] },
      { blockId: 'b-body', alignment: null, runs: [{ text: '本合同为固定期限劳动合同。' }] },
    ],
  };
  const doc = () => content([
    { id: 'b-title', kind: 'paragraph', text: '劳动合同书' },
    { id: 'b-clause', kind: 'paragraph', text: '第一条　合同期限' },
    { id: 'b-no', kind: 'paragraph', text: '编号：＿＿＿' },
    { id: 'b-body', kind: 'paragraph', text: '本合同为固定期限劳动合同。' },
  ]);

  it('turns an entirely bold paragraph into a heading and keeps alignment', () => {
    const { plan, report } = planFromContent(doc(), {}, FORMATTING);
    const blocks = docOf(plan).blocks;
    expect(blocks.map((block) => block.style)).toEqual(['Title', 'Heading1', 'Normal', 'Normal']);
    expect(blocks.map((block) => (block as { alignment?: string }).alignment)).toEqual(['center', undefined, 'right', undefined]);
    expect(report.carried.boldAsHeading).toBe(2);
    expect(report.carried.aligned).toBe(2);
    expect(report.notes.join(' ')).toMatch(/Presentational facts read/u);
  });

  it('carries nothing presentational when the observation was not requested', () => {
    const { plan, report } = planFromContent(doc(), {});
    const blocks = docOf(plan).blocks;
    expect(blocks.every((block) => block.style === 'Normal')).toBe(true);
    expect(blocks.every((block) => (block as { alignment?: string }).alignment === undefined)).toBe(true);
    expect(report.carried.boldAsHeading).toBe(0);
    expect(report.notes.join(' ')).toMatch(/No presentational observation was requested/u);
  });

  it('does not invent a heading from a paragraph that is only partly bold', () => {
    const mixed = { paragraphs: [{ blockId: 'b-1', alignment: null, runs: [{ text: '加粗', bold: true }, { text: '不加粗' }] }] };
    const { plan } = planFromContent(content([{ id: 'b-1', kind: 'paragraph', text: '加粗不加粗' }]), {}, mixed);
    expect(docOf(plan).blocks[0]!.style).toBe('Normal');
  });

  it('carries a declared run size and East Asian face into the rebuilt run', () => {
    const formatted = {
      paragraphs: [
        { blockId: 'b-no', alignment: 'right', runs: [{ text: '编号：＿＿＿', size: 10.5, eastAsia: '宋体' }] },
        { blockId: 'b-appendix', alignment: null, runs: [{ text: '附件：变更记录', size: 14, eastAsia: '黑体' }] },
      ],
    };
    const { plan, report } = planFromContent(content([
      { id: 'b-no', kind: 'paragraph', text: '编号：＿＿＿' },
      { id: 'b-appendix', kind: 'paragraph', text: '附件：变更记录' },
    ]), {}, formatted);
    const blocks = docOf(plan).blocks as unknown as { runs: { text: string; size?: number; eastAsia?: string }[] }[];
    expect(blocks[0]!.runs[0]).toEqual({ text: '编号：＿＿＿', size: 10.5, eastAsia: '宋体' });
    expect(blocks[1]!.runs[0]).toEqual({ text: '附件：变更记录', size: 14, eastAsia: '黑体' });
    expect(report.carried.sized).toBe(2);
    expect(report.carried.faced).toBe(2);
    // A size the reference never declared is never invented: the preset decides.
    const plain = planFromContent(content([{ id: 'b-1', kind: 'paragraph', text: '正文' }]), {}, { paragraphs: [{ blockId: 'b-1', alignment: null, runs: [{ text: '正文' }] }] });
    expect((docOf(plain.plan).blocks[0] as unknown as { runs: unknown[] }).runs[0]).toEqual({ text: '正文' });
  });

  it('keeps the dominant run when a paragraph declares conflicting sizes, and says so', () => {
    const mixedSizes = {
      paragraphs: [{ blockId: 'b-1', alignment: null, runs: [
        { text: '短', size: 22, eastAsia: '黑体' },
        { text: '这是一段更长的正文，尺寸应当是它说了算。', size: 12, eastAsia: '宋体' },
      ] }],
    };
    const { plan, report } = planFromContent(content([{ id: 'b-1', kind: 'paragraph', text: '短这是一段更长的正文，尺寸应当是它说了算。' }]), {}, mixedSizes);
    const run = (docOf(plan).blocks[0] as unknown as { runs: { size?: number; eastAsia?: string }[] }).runs[0]!;
    expect(run.size).toBe(12);
    expect(run.eastAsia).toBe('宋体');
    expect(report.notes.join(' ')).toMatch(/dominant run/u);
  });

  it('carries no colour: the register owns the palette', () => {
    const coloured = { paragraphs: [{ blockId: 'b-1', alignment: null, runs: [{ text: '红色正文', size: 12 }] }] };
    const { plan } = planFromContent(content([{ id: 'b-1', kind: 'paragraph', text: '红色正文' }]), {}, coloured);
    expect((docOf(plan).blocks[0] as unknown as { runs: Record<string, unknown>[] }).runs[0]).toEqual({ text: '红色正文', size: 12 });
  });

  const TWO_TABLES = content([
    { id: 't1', kind: 'table', gridColumns: 2, rows: [
      { cells: [{ text: '甲方（盖章）：' }, { text: '乙方（签名）：' }] },
      { cells: [{ text: '签订日期：' }, { text: '签订日期：' }] },
    ] },
    { id: 't2', kind: 'table', gridColumns: 2, rows: [
      { cells: [{ text: '变更日期' }, { text: '变更条款' }] },
      { cells: [{ text: '2026-09-26' }, { text: '工资调整' }] },
    ] },
  ]);
  const pointer = (tbl: number, tr: number, tc: number) => `word/document.xml!/w:document/w:body/w:tbl[${tbl}]/w:tr[${tr}]/w:tc[${tc}]/w:p[1]`;

  it('carries a cell’s own size, face and emphasis from the observed pointer', () => {
    const observed = { paragraphs: [
      { blockId: 'a', pointer: pointer(1, 1, 1), alignment: null, runs: [{ text: '甲方（盖章）：', bold: false, size: 12, eastAsia: '宋体' }] },
      { blockId: 'b', pointer: pointer(2, 1, 1), alignment: null, runs: [{ text: '变更日期', bold: true, size: 10.5, eastAsia: '黑体' }] },
      { blockId: 'c', pointer: pointer(2, 1, 2), alignment: null, runs: [{ text: '变更条款', bold: true, size: 10.5, eastAsia: '黑体' }] },
    ] };
    const { plan, report } = planFromContent(TWO_TABLES, {}, observed);
    const tables = docOf(plan).blocks.filter((block) => block.kind === 'table') as unknown as { cellFormats?: unknown }[];
    expect(tables[0]!.cellFormats).toEqual([[{ size: 12, eastAsia: '宋体', bold: false }, null], [null, null]]);
    expect(tables[1]!.cellFormats).toEqual([[{ size: 10.5, eastAsia: '黑体', bold: true }, { size: 10.5, eastAsia: '黑体', bold: true }], [null, null]]);
    expect(report.carried.cellFormats).toBe(3);
    expect(report.notes.join(' ')).toMatch(/table cell\(s\) kept their own size and face/u);
  });

  it('attaches cell formatting to the table the pointer names, not to the first one', () => {
    const observed = { paragraphs: [
      { blockId: 'b', pointer: pointer(2, 1, 1), alignment: null, runs: [{ text: '变更日期', size: 10.5 }] },
    ] };
    const { plan } = planFromContent(TWO_TABLES, {}, observed);
    const tables = docOf(plan).blocks.filter((block) => block.kind === 'table') as unknown as { cellFormats?: unknown }[];
    expect(tables[0]!.cellFormats).toBeUndefined();
    expect((tables[1]!.cellFormats as { size?: number }[][] | undefined)?.[0]?.[0]).toEqual({ size: 10.5 });
  });

  it('leaves the plan free of cell formatting when the observation has none', () => {
    const observed = { paragraphs: [{ blockId: 'p', pointer: 'word/document.xml!/w:document/w:body/w:p[1]', alignment: null, runs: [{ text: '正文' }] }] };
    const { plan, report } = planFromContent(TWO_TABLES, {}, observed);
    const tables = docOf(plan).blocks.filter((block) => block.kind === 'table') as unknown as { cellFormats?: unknown }[];
    expect(tables.every((table) => table.cellFormats === undefined)).toBe(true);
    expect(report.carried.cellFormats).toBe(0);
    expect(report.notes.join(' ')).not.toMatch(/table cell\(s\) kept/u);
  });

  it('carries the reference’s first-line indent, including the ones that are zero', () => {
    const observed = { paragraphs: [
      // 480 twips is two characters of 12 pt body text.
      { blockId: 'b-body', pointer: 'word/document.xml!/w:document/w:body/w:p[1]', alignment: null, runs: [{ text: '根据《中华人民共和国劳动法》', size: 12 }], indent: { firstLineTwips: 480 } },
      // A signature line: the reference states no indent at all, and the style
      // chain confirms it, so the rebuild must state zero rather than inherit.
      { blockId: 'b-sign', pointer: 'word/document.xml!/w:document/w:body/w:p[2]', alignment: null, runs: [{ text: '甲方（盖章）：', size: 12 }], indent: { firstLineTwips: 0, firstLineChars: 0 } },
      // Character-form indent, against a 10.5 pt run: still the same 24 pt.
      { blockId: 'b-note', pointer: 'word/document.xml!/w:document/w:body/w:p[3]', alignment: null, runs: [{ text: '注：本合同为示范文本', size: 10.5 }], indent: { firstLineChars: 400 } },
    ] };
    const doc = content([
      { id: 'b-body', kind: 'paragraph', text: '根据《中华人民共和国劳动法》' },
      { id: 'b-sign', kind: 'paragraph', text: '甲方（盖章）：' },
      { id: 'b-note', kind: 'paragraph', text: '注：本合同为示范文本' },
    ]);
    const { plan, report } = planFromContent(doc, {}, observed);
    const blocks = docOf(plan).blocks as unknown as { firstLineIndentPt?: number }[];
    expect(blocks.map((block) => block.firstLineIndentPt)).toEqual([24, 0, 42]);
    expect(report.carried.indented).toBe(3);
    expect(report.notes.join(' ')).toMatch(/first-line indent/u);
  });

  it('states no indent when the observation has none, rather than guessing zero', () => {
    const observed = { paragraphs: [{ blockId: 'b-1', pointer: 'word/document.xml!/w:document/w:body/w:p[1]', alignment: null, runs: [{ text: '正文' }] }] };
    const { plan, report } = planFromContent(content([{ id: 'b-1', kind: 'paragraph', text: '正文' }]), {}, observed);
    expect((docOf(plan).blocks[0] as unknown as { firstLineIndentPt?: number }).firstLineIndentPt).toBeUndefined();
    expect(report.carried.indented).toBe(0);
  });

  it('reads the effective first-line indent out of a real document', async () => {
    const c = await context();
    try {
      // The preset's Normal style carries a two-character indent of its own, so
      // the second paragraph is stating that it wants none — the case that a
      // naive "read the paragraph's own w:ind" observation would get wrong.
      const created = await c.profile.call('docx-create', 'execute', {
        requestId: 'indent-source', operation: 'execute',
        plan: { kind: 'create', document: { preset: 'chinese-contract', blocks: [
          { kind: 'paragraph', id: 'body', runs: [{ text: '正文段落' }], firstLineIndentPt: 24 },
          { kind: 'paragraph', id: 'sign', runs: [{ text: '甲方（盖章）：' }], firstLineIndentPt: 0 },
        ] } },
      }) as { result: { artifactRef: unknown } };

      const on = await c.profile.call('docx-parse', 'execute', {
        artifactRef: created.result.artifactRef, requestId: 'indent-on', options: { featureFlags: { parseFormatting: true } },
      }) as { result: { ir: { content: { semantic: { blocks: unknown[] }; formatting?: { paragraphs: { runs: { text: string }[]; indent?: { firstLineTwips?: number; firstLineChars?: number } }[] } } } } };
      const observed = on.result.ir.content.formatting;
      expect(observed).toBeDefined();
      const indentOf = (text: string) => observed!.paragraphs.find((entry) => entry.runs.map((run) => run.text).join('') === text)?.indent;
      expect(indentOf('正文段落')).toEqual({ firstLineTwips: 480, firstLineChars: 200 });
      expect(indentOf('甲方（盖章）：')).toEqual({ firstLineTwips: 0, firstLineChars: 0 });

      // Planning from the same parse is the point: the observation is indexed by
      // semantic block id, so a hand-written fixture could not join to it.
      const { plan, report } = planFromContent(
        { semantic: on.result.ir.content.semantic as never },
        {},
        observed as never,
      );
      const blocks = docOf(plan).blocks as unknown as { firstLineIndentPt?: number }[];
      expect(blocks.map((block) => block.firstLineIndentPt)).toEqual([24, 0]);
      expect(report.carried.indented).toBe(2);
    } finally {
      await c.close();
    }
  });

  it('reports formatting from a real parse only when the parser was asked for it', async () => {    const c = await context();
    try {
      const created = await c.profile.call('docx-create', 'execute', {
        requestId: 'fmt-source', operation: 'execute',
        plan: { kind: 'create', document: { preset: 'chinese-contract', blocks: [
          // `bold` here is direct run formatting; the Title style's own bold and
          // size are inherited and therefore not part of the observation.
          { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '劳动合同书', bold: true }] },
          { kind: 'paragraph', id: 'no', runs: [{ text: '编号：＿＿＿' }], alignment: 'right' },
          { kind: 'paragraph', id: 'body', runs: [{ text: '本合同为固定期限劳动合同。' }] },
        ] } },
      }) as { result: { artifactRef: unknown } };
      const ref = created.result.artifactRef;

      const off = await c.profile.call('docx-parse', 'execute', { artifactRef: ref, requestId: 'fmt-off' }) as { result: { ir: { content: { formatting?: unknown } } } };
      expect(off.result.ir.content.formatting).toBeUndefined();

      const on = await c.profile.call('docx-parse', 'execute', {
        artifactRef: ref, requestId: 'fmt-on', options: { featureFlags: { parseFormatting: true } },
      }) as { result: { ir: { content: { formatting?: { paragraphs: { alignment: string | null; runs: { bold?: boolean; size?: number; eastAsia?: string }[] }[] } } } } };
      const observed = on.result.ir.content.formatting;
      expect(observed).toBeDefined();
      expect(observed!.paragraphs.length).toBe(3);
      // The observation records what a paragraph declares itself, not what it
      // inherits: the right-aligned paragraph carries `w:jc` directly, while the
      // Title paragraph is centred by its style and reports no alignment.
      expect(observed!.paragraphs.some((entry) => entry.alignment === 'right')).toBe(true);
      const boldRun = observed!.paragraphs.find((entry) => entry.runs.some((run) => run.bold === true));
      expect(boldRun).toBeDefined();
      // Faces and sizes are inherited from the style here, so the observation
      // does not carry them; `eastAsia` and `size` are covered by the unit cases
      // above and by a reference whose runs declare them.
    } finally {
      await c.close();
    }
  });
});

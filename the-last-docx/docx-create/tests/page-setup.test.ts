import { describe, expect, it } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import { styleXml } from '../src/engine/presets';
import { REGISTERS } from '../src/engine/design';
import type { DocumentSpec } from '../src/domain/docx-create';
import { memoryStore, part } from './helpers';

const BODY: DocumentSpec['blocks'] = [
  { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '劳动合同书' }] },
  { kind: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: '第一条　合同期限' }] },
  { kind: 'paragraph', id: 'p1', runs: [{ text: '本合同为固定期限劳动合同。' }] },
];

async function create(document: Partial<DocumentSpec>) {
  const memory = memoryStore();
  const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute({
    requestId: 'page-case', operation: 'execute', plan: { kind: 'create', document: { blocks: BODY, ...document } },
  });
  const bytes = memory.files.get(out.result.artifactRef.id)!;
  const optional = (name: string) => { try { return part(bytes, name); } catch { return undefined; } };
  return { out, body: part(bytes), footer: optional('word/footer1.xml'), header: optional('word/header1.xml') };
}

/** Twips for a millimetre value, the conversion the engine performs. */
const mm = (value: number) => Math.round(value * 1440 / 25.4);

describe('page geometry is the plan’s to state', () => {
  it('keeps the historical A4 with 25 mm all round when the plan says nothing', async () => {
    const { body } = await create({});
    expect(body).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(body).toContain('<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>');
  });

  it('sets a per-edge margin, which is what a bound document needs', async () => {
    // The approved sample: 2.5 cm top and bottom, 2.8 cm outer, 3.0 cm gutter.
    const { body } = await create({ page: { marginsMm: { top: 25, right: 28, bottom: 25, left: 30 } } });
    expect(body).toContain(`<w:pgMar w:top="${mm(25)}" w:right="${mm(28)}" w:bottom="${mm(25)}" w:left="${mm(30)}"`);
    expect(mm(30)).toBe(1701);
    expect(mm(28)).toBe(1587);
  });

  it('changes the printable width with the margins, so a table cannot overflow', async () => {
    const wide: DocumentSpec['blocks'] = [
      { kind: 'table', id: 't', rows: [['a', 'b']], columnWidthsMm: [80, 80] },
    ];
    // 160 mm fits the default 25 mm margins...
    const ok = await create({ blocks: wide });
    expect(ok.out.verification?.ok).toBe(true);
    // ...but not a 30 + 28 mm gutter, which leaves 152 mm.
    await expect(create({ blocks: wide, page: { marginsMm: { left: 30, right: 28 } } }))
      .rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
});

describe('the footer states where the reader is and how much is left', () => {
  it('renders 第 X 页 共 Y 页 from two fields', async () => {
    const { footer } = await create({ pageNumberStyle: 'pageOfTotal' });
    expect(footer).toContain('<w:fldSimple w:instr="PAGE">');
    expect(footer).toContain('<w:fldSimple w:instr="NUMPAGES">');
    expect(footer).toContain('第 ');
    expect(footer).toContain(' 页 共 ');
    expect(footer).toContain(' 页');
  });

  it('stays a lone PAGE field for a plan that does not ask for the total', async () => {
    const { footer } = await create({});
    expect(footer).toContain('<w:fldSimple w:instr="PAGE">');
    expect(footer).not.toContain('NUMPAGES');
  });
});

describe('a layout table can state that it has no edges', () => {
  it('writes all six edges as none for a signature block', async () => {
    const { body } = await create({
      blocks: [{ kind: 'table', id: 'sign', borders: 'none', columnWidthsMm: [75, 75], rows: [['甲方（盖章）：', '乙方（签名）：']] }],
    });
    for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      expect(body).toContain(`<w:${side} w:val="none"/>`);
    }
    expect(body).not.toContain('w:val="single"');
  });

  it('still uses the register scheme when the table asks for nothing', async () => {
    const { body } = await create({
      scenario: 'formal-record',
      blocks: [{ kind: 'table', id: 'grid', columnWidthsMm: [75, 75], rows: [['a', 'b']] }],
    });
    expect(body).toContain('<w:insideV w:val="single" w:sz="4" w:color="000000"/>');
  });
});

describe('the contract preset follows the approved sample', () => {
  const styles = styleXml('chinese-contract', REGISTERS.plain);

  it('sets the clause heading at body size, bold, in 黑体', () => {
    const h1 = /<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/u.exec(styles)?.[0] ?? '';
    expect(h1).toContain('w:val="24"');            // 小四, same size as the body
    expect(h1).toContain('<w:b/>');
    expect(h1).toContain('w:eastAsia="SimHei"');
    expect(h1).toContain('<w:color w:val="000000"/>');
    expect(h1).toContain('w:before="200"');        // 10 pt above the clause
    expect(h1).toContain('w:after="80"');          // 4 pt below the heading
  });

  it('keeps the body at 小四 宋体 with a two-character indent and 1.5 leading', () => {
    expect(styles).toContain('<w:sz w:val="24"/>');
    expect(styles).toContain('w:line="360" w:lineRule="auto"');
    expect(styles).toContain('w:firstLineChars="200"');
    expect(styles).toContain('w:after="40"');      // 2 pt between paragraphs
  });

  it('does not change the report or long-form presets', () => {
    // A paper's first level stays 三号 黑体 unbolded; only the contract preset
    // sits at body size.
    const longForm = styleXml('chinese-long', REGISTERS.academic);
    const h1 = /<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/u.exec(longForm)?.[0] ?? '';
    expect(h1).toContain('w:val="32"');
    expect(h1).not.toContain('<w:b/>');
    const report = styleXml('report', REGISTERS.report);
    expect(report).toContain('w:after="120"');
  });
});

import { describe, expect, it } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import type { DocumentSpec } from '../src/domain/docx-create';
import { memoryStore, part } from './helpers';

async function create(document: Partial<DocumentSpec>) {
  const memory = memoryStore();
  const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute({
    requestId: 'cell-format', operation: 'execute', plan: { kind: 'create', document: { blocks: document.blocks ?? [], ...document } },
  });
  const bytes = memory.files.get(out.result.artifactRef.id)!;
  return { out, bytes, body: part(bytes) };
}

/** The `w:rPr` of the paragraph whose text is exactly this. */
function runPropertiesOf(body: string, text: string): string {
  const paragraph = body.split('</w:p>').find((piece) => piece.includes(`>${text}</w:t>`));
  if (paragraph === undefined) throw new Error(`no paragraph carrying ${text}`);
  return /<w:rPr>[\s\S]*?<\/w:rPr>/u.exec(paragraph)?.[0] ?? '';
}

const TABLE: DocumentSpec['blocks'] = [
  {
    kind: 'table', id: 't',
    rows: [['变更日期', '变更内容'], ['2026-09-26', '试用期工资调整']],
    columnWidthsMm: [70, 82],
    cellFormats: [[{ size: 10.5, eastAsia: '黑体' }, { size: 10.5, eastAsia: '宋体' }], [null, null]],
  },
];

describe('a cell may state its own size and face', () => {
  it('writes them into that cell only, in schema order', async () => {
    const { body } = await create({ blocks: TABLE });
    // 10.5 pt is 21 half-points, and CT_RPr orders rFonts before b before sz.
    const head = runPropertiesOf(body, '变更日期');
    expect(head).toContain('<w:rFonts w:eastAsia="黑体"/>');
    expect(head).toContain('<w:sz w:val="21"/><w:szCs w:val="21"/>');
    expect(head.indexOf('<w:rFonts')).toBeLessThan(head.indexOf('<w:sz '));
    // The cells that state nothing keep the register's table scheme.
    expect(runPropertiesOf(body, '试用期工资调整')).toBe('');
  });

  it('keeps the header row bold, with the stated size after the toggle', async () => {
    const { body } = await create({ blocks: TABLE });
    const head = runPropertiesOf(body, '变更日期');
    expect(head).toContain('<w:b/>');
    expect(head.indexOf('<w:b/>')).toBeLessThan(head.indexOf('<w:sz '));
  });

  it('leaves a table that states no cell formatting exactly as it was', async () => {
    const plain: DocumentSpec['blocks'] = [{
      kind: 'table', id: 't', rows: [['变更日期', '变更内容'], ['a', 'b']], columnWidthsMm: [70, 82],
    }];
    const withField: DocumentSpec['blocks'] = [{
      kind: 'table', id: 't', rows: [['变更日期', '变更内容'], ['a', 'b']], columnWidthsMm: [70, 82],
      cellFormats: [[null, null], [null, null]],
    }];
    const before = await create({ blocks: plain });
    const after = await create({ blocks: withField });
    expect(after.body).toBe(before.body);
    expect(after.body).not.toContain('<w:sz w:val="21"/>');
  });

  it('rejects a per-cell array that addresses a cell which does not exist', async () => {
    await expect(create({
      blocks: [{
        kind: 'table', id: 't', rows: [['a', 'b']], columnWidthsMm: [80, 80],
        cellFormats: [[{ size: 10.5 }, { size: 10.5 }, { size: 10.5 }]],
      }],
    })).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('lets a cell that states its own emphasis override the first-row default', async () => {
    // A form's first row is a heading by default; a reference that says its
    // signature labels are not bold has looked, and is believed.
    const notBold: DocumentSpec['blocks'] = [{
      kind: 'table', id: 't', rows: [['甲方（盖章）：', '乙方（签名）：']], columnWidthsMm: [76, 76],
      cellFormats: [[{ bold: false, size: 12, eastAsia: '宋体' }, { bold: false }]],
    }];
    const { body } = await create({ blocks: notBold });
    const run = runPropertiesOf(body, '甲方（盖章）：');
    expect(run).not.toContain('<w:b/>');
    expect(run).toContain('<w:rFonts w:eastAsia="宋体"/>');
    expect(run).toContain('<w:sz w:val="24"/>');
  });
});

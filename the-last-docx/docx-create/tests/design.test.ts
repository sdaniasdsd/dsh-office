import { describe, expect, it } from 'vitest';
import { createDocxCreateModule } from '../src/index';
import { resolveDesign, SCENARIO_RULES, REGISTERS } from '../src/engine/design';
import { styleXml } from '../src/engine/presets';
import { scenarioSchema } from '../src/domain/docx-create';
import type { DocumentSpec } from '../src/domain/docx-create';
import { memoryStore, part } from './helpers';

const BLOCKS: DocumentSpec['blocks'] = [
  { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '第三季度客户交付简报' }] },
  { kind: 'paragraph', id: 'summary', style: 'Heading1', runs: [{ text: '汇总' }] },
  {
    kind: 'table', id: 'totals', header: true, columnWidthsMm: [38, 22, 100],
    rows: [['指标', '数量', '说明'], ['交付任务', '18', '本季度完成']],
  },
];

async function create(document: Partial<DocumentSpec>) {
  const memory = memoryStore();
  const module = createDocxCreateModule({ artifactStore: memory.store });
  const out = await module.handlers.execute({
    requestId: 'design-case', operation: 'execute', plan: { kind: 'create', document: { blocks: BLOCKS, ...document } },
  });
  const bytes = memory.files.get(out.result.artifactRef.id)!;
  return { out, body: part(bytes), styles: part(bytes, 'word/styles.xml') };
}

describe('decoration register', () => {
  it('judges every scenario the schema accepts', () => {
    for (const scenario of scenarioSchema.options) expect(SCENARIO_RULES[scenario]).toBeDefined();
  });

  it('decides how much decoration a scenario warrants, and says why', () => {
    const expected: Record<string, [string, string]> = {
      'internal-review': ['report', 'moderate'],
      'client-delivery': ['report', 'moderate'],
      'academic-report': ['academic', 'restrained'],
      'formal-record': ['plain', 'none'],
      'technical-spec': ['grid', 'restrained'],
    };
    for (const [scenario, [register, decoration]] of Object.entries(expected)) {
      const decision = resolveDesign({ preset: 'report', scenario: scenario as keyof typeof SCENARIO_RULES });
      expect(decision.register).toBe(register);
      expect(decision.decoration).toBe(decoration);
      expect(decision.source).toBe('scenario');
      // A judgement nobody can read is a judgement nobody can dispute.
      expect(decision.reason.length).toBeGreaterThan(20);
    }
  });

  it('lets a caller overrule the judgement explicitly', () => {
    const decision = resolveDesign({ preset: 'report', scenario: 'formal-record', register: 'grid' });
    expect(decision.register).toBe('grid');
    expect(decision.source).toBe('explicit');
    expect(decision.scenario).toBe('formal-record');
  });

  it('declares a compatibility default rather than inventing a scenario', () => {
    const decision = resolveDesign({ preset: 'technical' });
    expect(decision.register).toBe('legacy');
    expect(decision.source).toBe('compatibility-default');
    expect(decision.scenario).toBeUndefined();
  });
});

describe('the register reaches the document', () => {
  it('a plan that declares nothing keeps the historical table markup', async () => {
    const { out, body } = await create({});
    expect(out.result.design?.source).toBe('compatibility-default');
    const borders = body.slice(body.indexOf('<w:tblBorders>'), body.indexOf('</w:tblBorders>') + 16);
    expect(borders).toContain('<w:top w:val="single" w:sz="4" w:color="D0D5DD"/>');
    expect(borders).toContain('<w:insideV w:val="single" w:sz="4" w:color="D0D5DD"/>');
    expect(body).toContain('<w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>');
    expect(body).not.toContain('<w:tblLook');
    expect(body).not.toContain('<w:vAlign');
    expect(body).toContain('<w:shd w:fill="24476B"/>');
  });

  it('an internal brief gets heading rules and drops the rules between columns', async () => {
    const { out, body, styles } = await create({ scenario: 'internal-review' });
    expect(out.result.design?.register).toBe('report');
    expect(body).toContain('<w:left w:val="none"/>');
    expect(body).toContain('<w:insideV w:val="none"/>');
    expect(body).not.toContain('<w:insideV w:val="single"');
    expect(body).toContain('<w:shd w:fill="24476B"/>');
    expect(body).toContain('<w:vAlign w:val="center"/>');
    // 0.75 pt rule, three points below the heading text. `w:space` is points,
    // not twentieths - carrying it through the twips helper asked for a
    // twentieth of a point and pinned the rule to the glyphs.
    expect(styles).toContain('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="3" w:color="D0D5DD"/></w:pBdr>');
  });

  it('an academic report gets a three-line table with nothing added', async () => {
    const { out, body, styles } = await create({ scenario: 'academic-report', preset: 'chinese-long' });
    expect(out.result.design?.register).toBe('academic');
    expect(out.result.design?.decoration).toBe('restrained');
    // Outer rules heavy, the single rule under the header lighter.
    expect(body).toContain('<w:top w:val="single" w:sz="12" w:color="000000"/>');
    expect(body).toContain('<w:bottom w:val="single" w:sz="12" w:color="000000"/>');
    expect(body).toContain('<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>');
    // Absence has to be stated: an omitted edge inherits the table style's rule.
    expect(body).toContain('<w:right w:val="none"/>');
    expect(body).toContain('<w:insideH w:val="none"/>');
    expect(body).toContain('<w:insideV w:val="none"/>');
    // White table throughout, and conditional formatting switched off.
    expect(body).not.toContain('<w:shd');
    expect(body).toContain('<w:tblLook w:val="0000"/>');
    // Table text one step below the body, and no decorative heading rule.
    expect(body).toContain('<w:sz w:val="23"/>');
    expect(styles).not.toContain('<w:pBdr>');
    // First-line indent in characters, so it scales with the body size.
    expect(styles).toContain('w:firstLineChars="200"');
  });

  it('a formal record is left undecorated', async () => {
    const { out, body, styles } = await create({ scenario: 'formal-record' });
    expect(out.result.design?.decoration).toBe('none');
    expect(body).not.toContain('<w:shd');
    expect(styles).not.toContain('<w:pBdr>');
    expect(body).not.toContain('<w:vAlign');
  });

  it('writes the pPr defaults in the sequence ECMA-376 declares', () => {
    // widowControl precedes spacing; the reverse makes Word offer to repair the file.
    const styles = styleXml('report');
    const widow = styles.indexOf('<w:widowControl/>');
    const spacing = styles.indexOf('<w:spacing w:after="120"');
    expect(widow).toBeGreaterThan(-1);
    expect(spacing).toBeGreaterThan(widow);
  });
});

describe('typography is the preset’s business, decoration is the register’s', () => {
  it('sets a formal Chinese document the way the convention requires', () => {
    // 二号 title centred, 三号/四号 黑体 headings left unbolded, 小四 bold 宋体
    // third level, 小四 宋体 body at 1.5 line with a two-character first-line
    // indent, and no colour anywhere: a formal contract is not a decorated
    // document, but it is still a *standard* one.
    const styles = styleXml('chinese-long', REGISTERS.plain);
    expect(styles).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/>');
    expect(styles).toContain('<w:sz w:val="44"/>');
    expect(styles).toContain('w:firstLineChars="200"');
    // Heading 1: 三号 (16 pt = 32 half-points), 黑体, unbolded, black.
    const h1 = /<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/u.exec(styles)?.[0] ?? '';
    expect(h1).toContain('w:val="32"');
    expect(h1).toContain('w:eastAsia="SimHei"');
    expect(h1).toContain('<w:color w:val="000000"/>');
    expect(h1).not.toContain('<w:b/>');
    // Heading 3: 小四 (12 pt = 24 half-points), 宋体, bold.
    const h3 = /<w:style [^>]*w:styleId="Heading3">.*?<\/w:style>/u.exec(styles)?.[0] ?? '';
    expect(h3).toContain('w:val="24"');
    expect(h3).toContain('w:eastAsia="SimSun"');
    expect(h3).toContain('<w:b/>');
    expect(styles).not.toContain('20334D');
    // A formal record adds nothing of its own on top: no rule, no fill.
    expect(styles).not.toContain('<w:pBdr>');
  });

  it('leaves the Latin presets exactly as they were', () => {
    const styles = styleXml('report', REGISTERS.report);
    const h1 = /<w:style [^>]*w:styleId="Heading1">.*?<\/w:style>/u.exec(styles)?.[0] ?? '';
    expect(h1).toContain('<w:b/>');
    expect(h1).toContain('<w:color w:val="24476B"/>');
    expect(styles).not.toContain('w:firstLineChars');
  });
});

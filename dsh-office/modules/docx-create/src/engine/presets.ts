import type { Preset } from '../domain/docx-create';
import { REGISTERS, eighths, twips, type DesignTokens } from './design';
import { NS, xml } from './xml';

/** A register may name the preset's accent instead of repeating its value. */
export const ACCENT = 'ACCENT';

interface TitleSpec {
  /** Half-points, the unit `w:sz` stores. */
  size: number;
  bold: boolean;
  color: string;
  center: boolean;
}

interface HeadingSpec {
  size: number[];
  bold: boolean[];
  /** East-Asian face per level: Chinese formal documents set headings in 黑体. */
  eastAsia: string[];
  color: string;
  /** Space before each level, in points. */
  before: number[];
  /** Space after a heading, in points. */
  after: number;
}

/**
 * Original, code-defined masters. No third-party branding or undocumented
 * template license.
 *
 * Typography lives here; decoration lives in the design register. The split is
 * deliberate — a two-character first-line indent and a clause heading set at body
 * size are conventions of the document *type*, not decoration levels, so they
 * have to survive the register that adds nothing.
 *
 * `chinese-contract` exists because one Chinese preset cannot serve both a
 * paper and a contract: a paper's first-level heading is 三号 黑体 unbolded,
 * while a contract's clause heading is 小四 黑体 bold, sitting at body size.
 */
export const PRESETS = {
  report: {
    label: '通用报告', latin: 'Calibri', eastAsia: 'Microsoft YaHei', accent: '24476B',
    bodySize: 22, line: 300, indent: 0, paragraphAfter: 6, footerColor: '667085',
    title: { size: 40, bold: true, color: ACCENT, center: false } as TitleSpec,
    headings: { size: [30, 26, 22], bold: [true, true, true], eastAsia: ['Microsoft YaHei', 'Microsoft YaHei', 'Microsoft YaHei'], color: ACCENT, before: [12.5, 11, 9.5], after: 6 } as HeadingSpec,
  },
  technical: {
    label: '技术文档', latin: 'Calibri', eastAsia: 'Microsoft YaHei', accent: '176B6B',
    bodySize: 21, line: 290, indent: 0, paragraphAfter: 6, footerColor: '667085',
    title: { size: 40, bold: true, color: ACCENT, center: false } as TitleSpec,
    headings: { size: [30, 26, 22], bold: [true, true, true], eastAsia: ['Microsoft YaHei', 'Microsoft YaHei', 'Microsoft YaHei'], color: ACCENT, before: [12.5, 11, 9.5], after: 6 } as HeadingSpec,
  },
  'chinese-long': {
    label: '中文正式文书', latin: 'Times New Roman', eastAsia: 'SimSun', accent: '20334D',
    bodySize: 24, line: 360, indent: 2, paragraphAfter: 6, footerColor: '667085',
    // 二号 title centred, 三号/四号 黑体 headings left unbolded, 小四 bold 宋体
    // third level, everything black: the conventional hierarchy of a formal
    // Chinese *report*.
    title: { size: 44, bold: true, color: '000000', center: true } as TitleSpec,
    headings: { size: [32, 28, 24], bold: [false, false, true], eastAsia: ['SimHei', 'SimHei', 'SimSun'], color: '000000', before: [12.5, 11, 9.5], after: 6 } as HeadingSpec,
  },
  'chinese-contract': {
    label: '中文合同', latin: 'Times New Roman', eastAsia: 'SimSun', accent: '000000',
    bodySize: 24, line: 360, indent: 2, paragraphAfter: 2, footerColor: '000000',
    // A contract numbers its clauses `第一条 …` at body size rather than
    // building a chapter hierarchy: the clause heading is 小四 黑体 bold, sits
    // 10 pt above its clause and 4 pt below itself, and the page is black
    // throughout. Everything here is measured off the approved sample.
    title: { size: 44, bold: true, color: '000000', center: true } as TitleSpec,
    headings: { size: [24, 24, 24], bold: [true, true, true], eastAsia: ['SimHei', 'SimHei', 'SimSun'], color: '000000', before: [10, 10, 10], after: 4 } as HeadingSpec,
  },
} as const;

/** The fill a register asks for, resolved against the preset that owns the accent. */
export function headFillColor(preset: Preset, design: DesignTokens): string | undefined {
  const fill = design.table.headFill;
  if (!fill) return undefined;
  return fill === ACCENT ? PRESETS[preset].accent : fill;
}

/** The colour running matter is set in, which a formal document keeps black. */
export function footerColorOf(preset: Preset): string {
  return PRESETS[preset].footerColor;
}

const colorOf = (value: string, accent: string) => (value === ACCENT ? accent : value);

/**
 * The styles part for a preset, decorated according to the resolved register.
 *
 * The register only decides decoration. Typography (faces, sizes, leading, the
 * Chinese first-line indent, paragraph rhythm) stays with the preset, so a
 * scenario can change how a document is dressed without silently changing how it
 * reads.
 */
export function styleXml(preset: Preset, design: DesignTokens = REGISTERS.legacy): string {
  const p = PRESETS[preset];
  const fonts = `<w:rFonts w:ascii="${p.latin}" w:hAnsi="${p.latin}" w:eastAsia="${p.eastAsia}"/>`;
  const style = (id: string, name: string, size: number, extra = '', pp = '', faces = fonts) => `<w:style w:type="paragraph" w:styleId="${id}"${id === 'Normal' ? ' w:default="1"' : ''}><w:name w:val="${xml(name)}"/>${id !== 'Normal' ? '<w:basedOn w:val="Normal"/>' : ''}<w:next w:val="Normal"/><w:qFormat/><w:pPr>${pp}</w:pPr><w:rPr>${faces}${extra}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

  // A heading rule sits between the text and whatever follows it. `w:space` is
  // ST_PointMeasure — points, not twentieths.
  const rule = design.heading.bottomRule;
  const headingRule = rule
    ? `<w:pBdr><w:bottom w:val="single" w:sz="${eighths(rule.size)}" w:space="${Math.round(rule.space)}" w:color="${rule.color}"/></w:pBdr>`
    : '';
  const headingAlignment = design.heading.alignment === 'center' ? '<w:jc w:val="center"/>' : '';
  // Chinese first-line indent is defined in characters so it scales with the
  // body size. `w:firstLineChars` is hundredths of a character; `w:firstLine`
  // carries the equivalent twips for readers that ignore the *Chars form.
  const indentChars = p.indent;
  const bodyIndent = indentChars > 0
    ? `<w:ind w:firstLineChars="${Math.round(indentChars * 100)}" w:firstLine="${Math.round(indentChars * p.bodySize * 10)}"/>`
    : '';
  // The body indent lives in `w:pPrDefault`, so every paragraph inherits it —
  // including headings, which a formal document sets flush left, and the title,
  // whose centring the indent silently shifts right by two characters. Reset it
  // wherever it does not belong, and only when there is an indent to reset.
  const noIndent = indentChars > 0 ? '<w:ind w:firstLineChars="0" w:firstLine="0"/>' : '';

  const titleExtra = `${p.title.bold ? '<w:b/>' : ''}<w:color w:val="${colorOf(p.title.color, p.accent)}"/>`;
  const titleParagraph = `<w:keepNext/><w:spacing w:before="320" w:after="240"/>${noIndent}${p.title.center ? '<w:jc w:val="center"/>' : ''}`;
  // A formal Chinese document sets its title in 黑体 too, so the title takes the
  // first heading's face rather than the body's. For the Latin presets that face
  // is the body face, so their output does not move.
  const headingFaces = (level: number) => `<w:rFonts w:ascii="${p.latin}" w:hAnsi="${p.latin}" w:eastAsia="${p.headings.eastAsia[level]}"/>`;

  // w:pPr children in the sequence ECMA-376 declares: widowControl precedes
  // spacing, spacing precedes ind, outlineLvl comes last.
  return `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr>${fonts}<w:sz w:val="${p.bodySize}"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/><w:spacing w:after="${twips(p.paragraphAfter)}" w:line="${p.line}" w:lineRule="auto"/>${bodyIndent}</w:pPr></w:pPrDefault></w:docDefaults>${
    style('Normal', 'Normal', p.bodySize, '', bodyIndent) +
    style('Title', 'Title', p.title.size, titleExtra, titleParagraph, headingFaces(0)) +
    style('Subtitle', 'Subtitle', 24, '<w:color w:val="667085"/>', `<w:keepNext/>${noIndent}`) +
    [0, 1, 2].map((level) => style(
      `Heading${level + 1}`, `heading ${level + 1}`, p.headings.size[level]!,
      `${p.headings.bold[level] ? '<w:b/>' : ''}<w:color w:val="${colorOf(p.headings.color, p.accent)}"/>`,
      `<w:keepNext/><w:keepLines/>${headingRule}<w:spacing w:before="${twips(p.headings.before[level]!)}" w:after="${twips(p.headings.after)}"/>${noIndent}${headingAlignment}<w:outlineLvl w:val="${level}"/>`,
      headingFaces(level),
    )).join('') +
    style('Caption', 'Caption', 18, '<w:color w:val="667085"/>', '<w:keepLines/>') +
    `<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:fill="F2F4F7"/><w:spacing w:line="260" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="${p.eastAsia}"/><w:sz w:val="19"/></w:rPr></w:style>`
  }</w:styles>`;
}

import type { DecorationLevel, DesignDecision, DesignRegister, Preset, Scenario } from '../domain/docx-create';

export type { DecorationLevel, DesignDecision };

/** ECMA-376 states border width in eighths of a point (`w:sz`). */
const eighths = (points: number): number => Math.round(points * 8);
/** ECMA-376 states lengths in twentieths of a point (`w:w`). */
const twips = (points: number): number => Math.round(points * 20);

export interface DesignEdge {
  /** `none` is a stated absence, which is not the same as an omitted element. */
  style: 'single' | 'none';
  /** Points. Ignored by renderers when `style` is `none`, but kept for the record. */
  size: number;
  color: string;
}

const NONE_EDGE: DesignEdge = { style: 'none', size: 0.5, color: 'auto' };
const edge = (size: number, color: string): DesignEdge => ({ style: 'single', size, color });

export interface TableDesign {
  borders: Record<'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV', DesignEdge>;
  /** Stated as a per-cell bottom border on the header row — the classic "栏目线". */
  headRule?: DesignEdge;
  headFill?: string;
  headTextColor?: string;
  /** Points. */
  cellMargins: { top: number; left: number; bottom: number; right: number };
  cellVerticalAlignment: 'top' | 'center' | 'bottom';
  /** Points added to the body size for text inside cells. Negative = one step down. */
  fontSizeDelta: number;
  /** `w:tblLook` conditional-formatting switches; `0000` disables all of them. */
  look?: string;
  /**
   * Line spacing inside cells, in 240ths of a line (240 = single).
   *
   * A form table is read row by row, so inheriting the body's 1.5-line setting
   * makes every row a quarter taller than it needs and turns a five-row
   * information block into a page of white space. Omitted by `legacy`, which is
   * a frozen compatibility floor rather than a design.
   */
  cellLineSpacing?: number;
}

export interface DesignTokens {
  label: string;
  intent: string;
  decoration: DecorationLevel;
  heading: { bottomRule?: { size: number; color: string; space: number }; alignment: 'left' | 'center' };
  table: TableDesign;
  /**
   * List indentation, in twips. Registers that say nothing use the standard
   * hanging indent; a document whose body carries a first-line indent needs its
   * markers to start further left so item text lines up with the body.
   */
  list?: { indent: { left: number; hanging: number } };
}

/**
 * Widths follow the OOXML three-line-table convention: rules at 1.5 pt for the
 * table's outer edges, 0.75 pt for the single rule between head and body, and
 * hairlines at 0.5 pt for interior separators. Interior rules must stay thinner
 * than the outer ones or the table reads as a grid.
 */
export const REGISTERS: Record<DesignRegister, DesignTokens> = {
  /**
   * The historical output, kept verbatim so a plan that declares nothing keeps
   * producing byte-identical XML. Not a design choice — a compatibility floor.
   */
  legacy: {
    label: '历史默认',
    intent: '未声明场景时的既有行为，仅用于向后兼容。',
    decoration: 'moderate',
    heading: { alignment: 'left' },
    table: {
      borders: { top: edge(0.5, 'D0D5DD'), left: edge(0.5, 'D0D5DD'), bottom: edge(0.5, 'D0D5DD'), right: edge(0.5, 'D0D5DD'), insideH: edge(0.5, 'D0D5DD'), insideV: edge(0.5, 'D0D5DD') },
      headFill: 'ACCENT',
      headTextColor: 'FFFFFF',
      cellMargins: { top: 4.5, left: 5, bottom: 4.5, right: 5 },
      cellVerticalAlignment: 'top',
      fontSizeDelta: 0,
    },
  },
  /**
   * Modern report: the read is sequential and the reader is skimming, so a rule
   * under headings and a filled header row carry the structure. Interior
   * vertical rules are dropped — they add ink without adding information.
   */
  report: {
    label: '现代报告',
    intent: '内部评审、对外交付这类需要快速定位的短文档。',
    decoration: 'moderate',
    heading: { bottomRule: { size: 0.75, color: 'D0D5DD', space: 3 }, alignment: 'left' },
    table: {
      borders: { top: edge(0.75, '9AA6B2'), left: NONE_EDGE, bottom: edge(0.75, '9AA6B2'), right: NONE_EDGE, insideH: edge(0.5, 'D0D5DD'), insideV: NONE_EDGE },
      headFill: 'ACCENT',
      headTextColor: 'FFFFFF',
      cellMargins: { top: 6, left: 8, bottom: 6, right: 8 },
      cellVerticalAlignment: 'center',
      fontSizeDelta: 0,
      look: '0000',
      cellLineSpacing: 240,
    },
  },
  /**
   * GB three-line table. Decoration here is a *specification*, not a taste:
   * white cells throughout (no banding, no header fill), no vertical rules, no
   * rule between body rows, outer rules heavier than the single head rule.
   */
  academic: {
    label: '学术三线表',
    intent: '论文、正式研究报告等受排版规范约束的文档。',
    decoration: 'restrained',
    heading: { alignment: 'left' },
    table: {
      borders: { top: edge(1.5, '000000'), left: NONE_EDGE, bottom: edge(1.5, '000000'), right: NONE_EDGE, insideH: NONE_EDGE, insideV: NONE_EDGE },
      headRule: edge(0.75, '000000'),
      cellMargins: { top: 3, left: 5, bottom: 3, right: 5 },
      cellVerticalAlignment: 'center',
      fontSizeDelta: -0.5,
      look: '0000',
      cellLineSpacing: 240,
    },
    // Chinese formal documents indent the body by two characters (the preset
    // carries that indent), so the markers hang further left than the default in
    // order to line item text up with the indented body rather than sit proud.
    list: { indent: { left: 480, hanging: 240 } },
  },
  /** Archival and evidentiary documents: no rules, no fills, nothing decorative. */
  plain: {
    label: '无装帧存档',
    intent: '合同、凭证、需长期保存与机器读取的存档文档。',
    decoration: 'none',
    heading: { alignment: 'left' },
    table: {
      borders: { top: edge(0.5, '000000'), left: edge(0.5, '000000'), bottom: edge(0.5, '000000'), right: edge(0.5, '000000'), insideH: edge(0.5, '000000'), insideV: edge(0.5, '000000') },
      cellMargins: { top: 4.5, left: 5, bottom: 4.5, right: 5 },
      cellVerticalAlignment: 'top',
      fontSizeDelta: -0.5,
      look: '0000',
      cellLineSpacing: 240,
    },
  },
  /**
   * Data-dense tables keep the full grid on purpose: column alignment is doing
   * informational work, and removing the vertical rules measurably slows lookup
   * across a wide row.
   */
  grid: {
    label: '数据网格',
    intent: '技术规格、参数表等列多且需要逐行对齐查阅的文档。',
    decoration: 'restrained',
    heading: { bottomRule: { size: 0.5, color: 'D0D5DD', space: 3 }, alignment: 'left' },
    table: {
      borders: { top: edge(0.5, '9AA6B2'), left: edge(0.5, '9AA6B2'), bottom: edge(0.5, '9AA6B2'), right: edge(0.5, '9AA6B2'), insideH: edge(0.5, 'D0D5DD'), insideV: edge(0.5, 'D0D5DD') },
      headFill: 'F2F4F7',
      headTextColor: '101828',
      cellMargins: { top: 4, left: 6, bottom: 4, right: 6 },
      cellVerticalAlignment: 'center',
      fontSizeDelta: -0.5,
      look: '0000',
      cellLineSpacing: 240,
    },
  },
};

interface ScenarioRule {
  register: DesignRegister;
  /** Spelled out so a reviewer can disagree with the judgement, not just its result. */
  reason: string;
}

const SCENARIOS: Record<Scenario, ScenarioRule> = {
  'internal-review': {
    register: 'report',
    reason: '内部评审由同组织读者快速浏览，标题分隔线与表头底纹承担定位功能；不采用学术规范，也不加品牌化装饰。',
  },
  'client-delivery': {
    register: 'report',
    reason: '对外交付物需要稳定的视觉层级体现完成度，但可读性优先，仍不引入装饰性元素。',
  },
  'academic-report': {
    register: 'academic',
    reason: '受排版规范约束：三线表、全表白底、无竖线是硬性要求，此处装帧是合规问题而非审美选择。',
  },
  'formal-record': {
    register: 'plain',
    reason: '存档与凭证以长期可读和机器解析为先，任何装饰性装帧都是缺陷，故不加标题线、底纹与字号变化。',
  },
  'technical-spec': {
    register: 'grid',
    reason: '列多且需要逐行比对，完整网格承担信息功能；去掉竖线会降低查阅效率，因此保留。',
  },
};

/**
 * Resolve the design register. An explicit `register` always wins, a declared
 * `scenario` decides otherwise, and a plan that declares neither keeps the
 * historical output. The decision and its reason travel back to the caller.
 */
export function resolveDesign(input: { preset: Preset; scenario?: Scenario; register?: DesignRegister }): DesignDecision {
  if (input.register) {
    const tokens = REGISTERS[input.register];
    return {
      register: input.register,
      label: tokens.label,
      decoration: tokens.decoration,
      intent: tokens.intent,
      source: 'explicit',
      reason: `调用方显式指定装帧寄存器 ${input.register}，覆盖场景推断。`,
      ...(input.scenario ? { scenario: input.scenario } : {}),
    };
  }
  if (input.scenario) {
    const rule = SCENARIOS[input.scenario];
    const tokens = REGISTERS[rule.register];
    return {
      register: rule.register,
      label: tokens.label,
      decoration: tokens.decoration,
      intent: tokens.intent,
      source: 'scenario',
      reason: rule.reason,
      scenario: input.scenario,
    };
  }
  const tokens = REGISTERS.legacy;
  return {
    register: 'legacy',
    label: tokens.label,
    decoration: tokens.decoration,
    intent: tokens.intent,
    source: 'compatibility-default',
    reason: `未声明 scenario 或 register，沿用 ${input.preset} 预设的既有输出以保证与历史行为一致。`,
  };
}

/** Every scenario this layer can judge, for callers that enumerate it. */
export const SCENARIO_RULES: Readonly<Record<Scenario, ScenarioRule>> = SCENARIOS;
export { eighths, twips };

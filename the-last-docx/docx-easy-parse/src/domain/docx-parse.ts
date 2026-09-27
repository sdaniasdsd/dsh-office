/**
 * docx-parse 的领域模型 —— 本模块【唯一】的文档数据结构。
 *
 * 这里定义的形状同时用于两处，因此不存在「双 IR」：
 *   1. Python 解析引擎通过 stdout 输出的 JSON 载荷，经本文件的校验器
 *      （信任边界）后成为 `ParseResult`；
 *   2. `mapper.ts` 在 `ParseResult` 之上补充 `FormatProfile`、
 *      派生 outline / counts，并整理顺序，直接得到公开的 `DocxParseIR`。
 *
 * 也就是说：引擎输出 = 领域模型 = IR 的数据部分（形状完全一致，不做二次搬运）。
 * mapper 只做「裁剪 → 判可用 → 补 profile → 派生/排序」，不重塑结构。
 *
 * 分层意图：
 *   engine(进程外) → ParseResult(本文件，中立) → mapper(裁决/补全/排序) → DocxParseIR
 */
import type { ArtifactRef, FormatProfile } from 'office-core';

/* -------------------------------------------------------------------------- */
/* 轻量枚举                                                                     */
/* -------------------------------------------------------------------------- */

export const CONTAINER_KINDS = ['zip', 'ole', 'rtf', 'unknown'] as const;
export type ContainerKind = (typeof CONTAINER_KINDS)[number];

/** 包所属的 OOXML 文档族。 */
export const DOCUMENT_KINDS = [
  'wordprocessingml',
  'spreadsheetml',
  'presentationml',
  'unknown',
] as const;
export type DocumentKind = (typeof DOCUMENT_KINDS)[number];

/** WordprocessingML 的四个变体（决定 artifact 的扩展名与媒体类型）。 */
export const DOCUMENT_VARIANTS = [
  'document',
  'template',
  'macroEnabled',
  'macroEnabledTemplate',
] as const;
export type DocumentVariant = (typeof DOCUMENT_VARIANTS)[number];

/** 段落对齐方式（来自 `w:pPr/w:jc/@w:val`）。 */
export const ALIGNMENTS = ['left', 'center', 'right', 'both', 'distribute', 'unknown'] as const;
export type Alignment = (typeof ALIGNMENTS)[number];

/** 样式种类（来自 `w:styles/w:style/@w:type`）。 */
export const STYLE_TYPES = ['paragraph', 'character', 'table', 'numbering', 'unknown'] as const;
export type StyleType = (typeof STYLE_TYPES)[number];

/* -------------------------------------------------------------------------- */
/* 段落 / 表格                                                                  */
/* -------------------------------------------------------------------------- */

/** 一个文本运行（run）及其基础字符格式。 */
export interface RunRecord {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
}

/**
 * 一个段落。
 *
 * `headingLevel` 是【派生】字段：由样式名（Heading N / 标题 N）或
 * `w:outlineLvl` 推导而来；无法判定时为 null。它让「标题」成为一等公民，
 * 而不必让每个消费者都去解析样式表。
 */
export interface ParagraphRecord {
  /** 在整篇文档中的段落序号（从 0 起，跨表格外的正文顺序）。 */
  index: number;
  /** 段落纯文本（拼接全部 `w:t`）。 */
  text: string;
  /** 段落样式 id（`w:pStyle/@w:val`），无则 null。 */
  styleId: string | null;
  /** 段落样式名（经 styles.xml 解析），无则 null。 */
  styleName: string | null;
  /** `w:outlineLvl` 原值（0 基），无则 null。 */
  outlineLevel: number | null;
  /** 标题层级（1 基）；非标题为 null。 */
  headingLevel: number | null;
  /** 对齐方式。 */
  alignment: Alignment | null;
  /** 是否属于编号/项目符号列表（存在 `w:numPr`）。 */
  listItem: boolean;
  /** 段落内各 run 的文本与格式。 */
  runs: RunRecord[];
}

/** 表格中的一个单元格（扁平记录，避免无界递归）。 */
export interface TableCell {
  row: number;
  column: number;
  text: string;
  paragraphCount: number;
  /** 横向合并跨度（`w:gridSpan`），默认 1。 */
  columnSpan: number;
  /** 纵向合并跨度（`w:vMerge`），默认 1；仅记录起始单元格。 */
  rowSpan: number;
}

/** 一张表格。 */
export interface TableRecord {
  /** 在整篇文档中的表格序号（从 0 起）。 */
  index: number;
  rows: number;
  columns: number;
  /** 表格样式 id（`w:tblPr/w:tblStyle/@w:val`），无则 null。 */
  styleId: string | null;
  cells: TableCell[];
}

/**
 * 文档正文中「按出现顺序」排列的块。
 *
 * 用判别联合而非两个独立数组：文档顺序本身是有意义的信息，
 * 拆成两个数组就会丢失「第 3 个块是表格」这类事实。
 */
export type Block =
  | { kind: 'paragraph'; paragraph: ParagraphRecord }
  | { kind: 'table'; table: TableRecord };

/* -------------------------------------------------------------------------- */
/* 样式 / 关系 / 批注 / 脚注                                                     */
/* -------------------------------------------------------------------------- */

/** 一条样式定义。 */
export interface StyleRecord {
  styleId: string;
  name: string | null;
  type: StyleType;
  /** 继承来源样式 id（`w:basedOn/@w:val`），无则 null。 */
  basedOn: string | null;
  /** 是否为该类型的默认样式（`w:default="1"`）。 */
  isDefault: boolean;
  /** 由样式名推导的标题层级（1 基），非标题为 null。 */
  headingLevel: number | null;
}

/** 一条 OPC 关系。 */
export interface RelationshipRecord {
  sourcePart: string;
  id: string;
  type: string;
  target: string;
  targetMode: 'Internal' | 'External';
}

/** 一条批注（`word/comments.xml`）。 */
export interface CommentRecord {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
}

/** 一条脚注或尾注。 */
export interface FootnoteRecord {
  id: string;
  kind: 'footnote' | 'endnote';
  text: string;
}

/** 核心属性（`docProps/core.xml`）。 */
export interface DocumentMetadata {
  title: string | null;
  creator: string | null;
  lastModifiedBy: string | null;
  created: string | null;
  modified: string | null;
  revision: string | null;
}

/* -------------------------------------------------------------------------- */
/* 派生视图                                                                     */
/* -------------------------------------------------------------------------- */

/** 大纲（标题）条目，由正文块派生。 */
export interface OutlineEntry {
  level: number;
  text: string;
  paragraphIndex: number;
}

/** 规模统计（由数据派生，不信任引擎自报）。 */
export interface DocumentCounts {
  blocks: number;
  paragraphs: number;
  tables: number;
  headings: number;
  styles: number;
  relationships: number;
  comments: number;
  footnotes: number;
  endnotes: number;
}

/* -------------------------------------------------------------------------- */
/* 引擎载荷（信任边界之外的形状）                                                */
/* -------------------------------------------------------------------------- */

/** 引擎上报的非致命问题。 */
export interface ParseIssue {
  code: string;
  message: string;
  path?: string;
}

/**
 * 解析结果：引擎与模块之间的唯一数据契约。
 *
 * 所有字段都必须可 JSON 序列化（这是跨进程传输的前提）。
 */
export interface ParseResult {
  /** 解析协议版本，便于引擎与模块独立演进时做兼容判断。 */
  parseVersion: number;
  /** 实际使用的 XML 后端（lxml / stdlib）。 */
  xmlBackend: string;
  container: ContainerKind;
  /** 依据文件名得到的扩展名。 */
  extension: string | null;
  sizeBytes: number;
  /** 文件字节的 SHA-256（小写十六进制）。 */
  sha256: string;
  /** 三态：true=已加密，false=未加密，null=无法判断。 */
  encrypted: boolean | null;
  documentKind: DocumentKind;
  documentVariant: DocumentVariant | null;
  /** 主文档部件对应的产物媒体类型。 */
  mediaType: string | null;
  metadata: DocumentMetadata;
  blocks: Block[];
  styles: StyleRecord[];
  relationships: RelationshipRecord[];
  comments: CommentRecord[];
  footnotes: FootnoteRecord[];
  /** 首个被触发的资源预算名；null 表示未超限。 */
  limitHit: string | null;
  /** 引擎上报的问题列表（非致命）。 */
  issues: ParseIssue[];
  /** 致命错误；非 null 时其余字段仅供参考。 */
  error: { kind: string; message: string } | null;
}

/** 公开的解析 IR：领域数据 + 格式档案（与 docx-inspect 的 FormatIR 同层）。 */
export interface DocxParseIR {
  profile: FormatProfile;
  metadata: DocumentMetadata;
  /** 由标题段落派生的大纲。 */
  outline: OutlineEntry[];
  /** 按文档顺序排列的正文块。 */
  blocks: Block[];
  styles: StyleRecord[];
  relationships: RelationshipRecord[];
  comments: CommentRecord[];
  footnotes: FootnoteRecord[];
  counts: DocumentCounts;
}

/* -------------------------------------------------------------------------- */
/* 派生纯函数                                                                   */
/* -------------------------------------------------------------------------- */

/** 判断一个段落是否为标题。 */
export function isHeading(paragraph: ParagraphRecord): boolean {
  return paragraph.headingLevel !== null;
}

/**
 * 由正文块派生出大纲。
 *
 * 只取标题段落，顺序即文档顺序；这是「单一数据源 = blocks」的体现，
 * 因此 outline 永远不会与正文漂移。
 */
export function buildOutline(blocks: Block[]): OutlineEntry[] {
  const outline: OutlineEntry[] = [];
  for (const block of blocks) {
    if (block.kind !== 'paragraph') continue;
    const { paragraph } = block;
    if (paragraph.headingLevel === null) continue;
    outline.push({
      level: paragraph.headingLevel,
      text: paragraph.text,
      paragraphIndex: paragraph.index,
    });
  }
  return outline;
}

/** 由领域数据派生规模统计。 */
export function computeCounts(result: {
  blocks: Block[];
  styles: StyleRecord[];
  relationships: RelationshipRecord[];
  comments: CommentRecord[];
  footnotes: FootnoteRecord[];
}): DocumentCounts {
  let paragraphs = 0;
  let tables = 0;
  let headings = 0;
  for (const block of result.blocks) {
    if (block.kind === 'paragraph') {
      paragraphs += 1;
      if (block.paragraph.headingLevel !== null) headings += 1;
    } else {
      tables += 1;
    }
  }
  let footnotes = 0;
  let endnotes = 0;
  for (const note of result.footnotes) {
    if (note.kind === 'endnote') endnotes += 1;
    else footnotes += 1;
  }
  return {
    blocks: result.blocks.length,
    paragraphs,
    tables,
    headings,
    styles: result.styles.length,
    relationships: result.relationships.length,
    comments: result.comments.length,
    footnotes,
    endnotes,
  };
}

/** 是否存在可见内容（至少一个非空段落或一张表）。 */
export function hasVisibleContent(blocks: Block[]): boolean {
  return blocks.some((block) =>
    block.kind === 'paragraph' ? block.paragraph.text.trim() !== '' : block.table.rows > 0,
  );
}

/* -------------------------------------------------------------------------- */
/* 信任边界：载荷校验                                                            */
/* -------------------------------------------------------------------------- */

/**
 * 当引擎输出的形状不符合 `ParseResult` 时抛出。
 *
 * 之所以单独定义而不复用 `DocxParseError`：
 * 它属于「协议层」错误，由 `adapter.ts` 捕获后翻译成 `ENGINE_PROTOCOL_ERROR`，
 * 从而让领域层不依赖模块的错误码定义（保持依赖方向单一）。
 */
export class ParseProtocolError extends Error {
  constructor(
    message: string,
    /** 出错位置，例如 `$.blocks[3].paragraph.text`。 */
    readonly pointer: string,
  ) {
    super(`${message} (at ${pointer})`);
    this.name = 'ParseProtocolError';
  }
}

/*
 * 以下断言函数是「信任边界」的实现。
 *
 * 引擎输出来自子进程，必须视为不可信输入：任何字段进入模块内部之前都要逐一校验，
 * 否则一个畸形载荷可能让后续逻辑在隐式假设下崩溃或产生错误结论。
 * 每个断言都带 JSON 指针，便于快速定位是引擎哪一段输出不符合协议。
 */

function asRecord(value: unknown, pointer: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new ParseProtocolError('expected object', pointer);
  }
  return value as Record<string, unknown>;
}

function asArray(value: unknown, pointer: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new ParseProtocolError('expected array', pointer);
  }
  return value;
}

function asString(value: unknown, pointer: string): string {
  if (typeof value !== 'string') {
    throw new ParseProtocolError('expected string', pointer);
  }
  return value;
}

function asOptionalString(value: unknown, pointer: string): string | null {
  if (value === null || value === undefined) return null;
  return asString(value, pointer);
}

function asNumber(value: unknown, pointer: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ParseProtocolError('expected finite number', pointer);
  }
  return value;
}

function asInteger(value: unknown, pointer: string): number {
  const raw = asNumber(value, pointer);
  if (!Number.isInteger(raw)) {
    throw new ParseProtocolError('expected integer', pointer);
  }
  return raw;
}

function asOptionalInteger(value: unknown, pointer: string): number | null {
  if (value === null || value === undefined) return null;
  return asInteger(value, pointer);
}

/** 严格布尔断言：只接受真正的 boolean，不接受 0/1 或字符串。 */
function asBoolean(value: unknown, pointer: string): boolean {
  if (typeof value !== 'boolean') {
    throw new ParseProtocolError('expected boolean', pointer);
  }
  return value;
}

function asOptionalBoolean(value: unknown, pointer: string): boolean | null {
  if (value === null || value === undefined) return null;
  return asBoolean(value, pointer);
}

/** 校验枚举取值。`allowed` 用 `as const` 数组提供，保证类型收窄。 */
function asEnum<T extends readonly string[]>(
  value: unknown,
  pointer: string,
  allowed: T,
): T[number] {
  const raw = asString(value, pointer);
  if (!(allowed as readonly string[]).includes(raw)) {
    throw new ParseProtocolError(`unexpected value "${raw}"`, pointer);
  }
  return raw as T[number];
}

function asOptionalEnum<T extends readonly string[]>(
  value: unknown,
  pointer: string,
  allowed: T,
): T[number] | null {
  if (value === null || value === undefined) return null;
  return asEnum(value, pointer, allowed);
}

function validateRun(value: unknown, pointer: string): RunRecord {
  const record = asRecord(value, pointer);
  return {
    text: asString(record.text, `${pointer}.text`),
    bold: asBoolean(record.bold, `${pointer}.bold`),
    italic: asBoolean(record.italic, `${pointer}.italic`),
    underline: asBoolean(record.underline, `${pointer}.underline`),
  };
}

function validateParagraph(value: unknown, pointer: string): ParagraphRecord {
  const record = asRecord(value, pointer);
  return {
    index: asInteger(record.index, `${pointer}.index`),
    text: asString(record.text, `${pointer}.text`),
    styleId: asOptionalString(record.styleId, `${pointer}.styleId`),
    styleName: asOptionalString(record.styleName, `${pointer}.styleName`),
    outlineLevel: asOptionalInteger(record.outlineLevel, `${pointer}.outlineLevel`),
    headingLevel: asOptionalInteger(record.headingLevel, `${pointer}.headingLevel`),
    alignment: asOptionalEnum(record.alignment, `${pointer}.alignment`, ALIGNMENTS),
    listItem: asBoolean(record.listItem, `${pointer}.listItem`),
    runs: asArray(record.runs, `${pointer}.runs`).map((entry, at) =>
      validateRun(entry, `${pointer}.runs[${at}]`),
    ),
  };
}

function validateTableCell(value: unknown, pointer: string): TableCell {
  const record = asRecord(value, pointer);
  return {
    row: asInteger(record.row, `${pointer}.row`),
    column: asInteger(record.column, `${pointer}.column`),
    text: asString(record.text, `${pointer}.text`),
    paragraphCount: asInteger(record.paragraphCount, `${pointer}.paragraphCount`),
    columnSpan: asInteger(record.columnSpan, `${pointer}.columnSpan`),
    rowSpan: asInteger(record.rowSpan, `${pointer}.rowSpan`),
  };
}

function validateTable(value: unknown, pointer: string): TableRecord {
  const record = asRecord(value, pointer);
  return {
    index: asInteger(record.index, `${pointer}.index`),
    rows: asInteger(record.rows, `${pointer}.rows`),
    columns: asInteger(record.columns, `${pointer}.columns`),
    styleId: asOptionalString(record.styleId, `${pointer}.styleId`),
    cells: asArray(record.cells, `${pointer}.cells`).map((entry, at) =>
      validateTableCell(entry, `${pointer}.cells[${at}]`),
    ),
  };
}

function validateBlock(value: unknown, pointer: string): Block {
  const record = asRecord(value, pointer);
  const kind = asEnum(record.kind, `${pointer}.kind`, ['paragraph', 'table'] as const);
  if (kind === 'paragraph') {
    return { kind, paragraph: validateParagraph(record.paragraph, `${pointer}.paragraph`) };
  }
  return { kind, table: validateTable(record.table, `${pointer}.table`) };
}

function validateStyle(value: unknown, pointer: string): StyleRecord {
  const record = asRecord(value, pointer);
  return {
    styleId: asString(record.styleId, `${pointer}.styleId`),
    name: asOptionalString(record.name, `${pointer}.name`),
    type: asEnum(record.type, `${pointer}.type`, STYLE_TYPES),
    basedOn: asOptionalString(record.basedOn, `${pointer}.basedOn`),
    isDefault: asBoolean(record.isDefault, `${pointer}.isDefault`),
    headingLevel: asOptionalInteger(record.headingLevel, `${pointer}.headingLevel`),
  };
}

function validateRelationship(value: unknown, pointer: string): RelationshipRecord {
  const record = asRecord(value, pointer);
  return {
    sourcePart: asString(record.sourcePart, `${pointer}.sourcePart`),
    id: asString(record.id, `${pointer}.id`),
    type: asString(record.type, `${pointer}.type`),
    target: asString(record.target, `${pointer}.target`),
    targetMode: asEnum(record.targetMode, `${pointer}.targetMode`, [
      'Internal',
      'External',
    ] as const),
  };
}

function validateComment(value: unknown, pointer: string): CommentRecord {
  const record = asRecord(value, pointer);
  return {
    id: asString(record.id, `${pointer}.id`),
    author: asOptionalString(record.author, `${pointer}.author`),
    initials: asOptionalString(record.initials, `${pointer}.initials`),
    date: asOptionalString(record.date, `${pointer}.date`),
    text: asString(record.text, `${pointer}.text`),
  };
}

function validateFootnote(value: unknown, pointer: string): FootnoteRecord {
  const record = asRecord(value, pointer);
  return {
    id: asString(record.id, `${pointer}.id`),
    kind: asEnum(record.kind, `${pointer}.kind`, ['footnote', 'endnote'] as const),
    text: asString(record.text, `${pointer}.text`),
  };
}

function validateMetadata(value: unknown, pointer: string): DocumentMetadata {
  const record = asRecord(value, pointer);
  return {
    title: asOptionalString(record.title, `${pointer}.title`),
    creator: asOptionalString(record.creator, `${pointer}.creator`),
    lastModifiedBy: asOptionalString(record.lastModifiedBy, `${pointer}.lastModifiedBy`),
    created: asOptionalString(record.created, `${pointer}.created`),
    modified: asOptionalString(record.modified, `${pointer}.modified`),
    revision: asOptionalString(record.revision, `${pointer}.revision`),
  };
}

function validateIssue(value: unknown, pointer: string): ParseIssue {
  const record = asRecord(value, pointer);
  const issue: ParseIssue = {
    code: asString(record.code, `${pointer}.code`),
    message: asString(record.message, `${pointer}.message`),
  };
  // path 为可选字段：仅在存在时才写入，避免出现 `path: undefined`。
  if (record.path !== undefined && record.path !== null) {
    issue.path = asString(record.path, `${pointer}.path`);
  }
  return issue;
}

/**
 * 把原始引擎载荷校验并归一化为 `ParseResult`。
 *
 * 这是整个模块的信任边界：载荷来自子进程，被视为不可信输入，
 * 每个字段在使用前都被检查，因此畸形输出只能变成受控的协议错误，
 * 而不会污染后续逻辑。
 *
 * @throws ParseProtocolError 当载荷形状不符合协议时
 */
export function parseParseResult(payload: unknown): ParseResult {
  const root = asRecord(payload, '$');

  const blocks = asArray(root.blocks, '$.blocks').map((entry, index) =>
    validateBlock(entry, `$.blocks[${index}]`),
  );

  const styles = asArray(root.styles, '$.styles').map((entry, index) =>
    validateStyle(entry, `$.styles[${index}]`),
  );

  const relationships = asArray(root.relationships, '$.relationships').map((entry, index) =>
    validateRelationship(entry, `$.relationships[${index}]`),
  );

  const comments = asArray(root.comments, '$.comments').map((entry, index) =>
    validateComment(entry, `$.comments[${index}]`),
  );

  const footnotes = asArray(root.footnotes, '$.footnotes').map((entry, index) =>
    validateFootnote(entry, `$.footnotes[${index}]`),
  );

  const issues = asArray(root.issues, '$.issues').map((entry, index) =>
    validateIssue(entry, `$.issues[${index}]`),
  );

  const rawError = root.error;
  const error =
    rawError === null || rawError === undefined
      ? null
      : (() => {
          const record = asRecord(rawError, '$.error');
          return {
            kind: asString(record.kind, '$.error.kind'),
            message: asString(record.message, '$.error.message'),
          };
        })();

  return {
    parseVersion: asNumber(root.parseVersion, '$.parseVersion'),
    xmlBackend: asString(root.xmlBackend, '$.xmlBackend'),
    container: asEnum(root.container, '$.container', CONTAINER_KINDS),
    extension: asOptionalString(root.extension, '$.extension'),
    sizeBytes: asNumber(root.sizeBytes, '$.sizeBytes'),
    sha256: asString(root.sha256, '$.sha256'),
    encrypted: asOptionalBoolean(root.encrypted, '$.encrypted'),
    documentKind: asEnum(root.documentKind, '$.documentKind', DOCUMENT_KINDS),
    documentVariant: asOptionalEnum(root.documentVariant, '$.documentVariant', DOCUMENT_VARIANTS),
    mediaType: asOptionalString(root.mediaType, '$.mediaType'),
    metadata: validateMetadata(root.metadata, '$.metadata'),
    blocks,
    styles,
    relationships,
    comments,
    footnotes,
    limitHit: asOptionalString(root.limitHit, '$.limitHit'),
    issues,
    error,
  };
}

/** 供 mapper 复用的类型别名。 */
export type ParseArtifactRef = ArtifactRef;

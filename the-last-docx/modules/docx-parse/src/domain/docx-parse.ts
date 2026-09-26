/**
 * docx-parse 的领域模型。
 *
 * `ParseResult` 是对产物的「与引擎无关」的原始描述：Python 桥接把它序列化成 JSON 输出，
 * 除此之外没有任何东西跨越进程边界。
 *
 * 关键分层意图：
 *
 *     engine(进程外) → ParseResult(本文件，中立) → sourcemap(身份) → mapper(双视图) → office-core IR
 *
 * 为什么把「原始观察」和「身份/视图」分开？
 *   - 引擎只被允许陈述它【看到】的事实（这个元素在哪、文本是什么、有没有 paraId）；
 *   - 「这个节点叫什么 id」「它是标题还是段落」「语义视图长什么样」全由 TS 侧决定。
 *   这样换引擎不会改变身份方案与 IR 形状——正是 spec「底层引擎可替换而不改变
 *   公开接口」这条要求的落点。
 */
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

/** 标题层级的来源。 */
export const HEADING_LEVEL_SOURCES = ['outlineLevel', 'styleName', 'inferred'] as const;
export type HeadingLevelSource = (typeof HEADING_LEVEL_SOURCES)[number];

/** 注释种类。 */
export const ANNOTATION_KINDS = ['comment', 'footnote', 'endnote'] as const;
export type AnnotationKind = (typeof ANNOTATION_KINDS)[number];

/** 包内部件记录。 */
export interface PartRecord {
  name: string;
  sizeBytes: number;
  compressedSize: number;
}

/** 段落的列表编号信息。 */
export interface RawListInfo {
  numId: string;
  level: number;
}

/**
 * 引擎观察到的原始段落。
 *
 * 所有字段都是「所见即所得」：引擎不判断它是不是标题、级别该算几级，
 * 只如实上报 pStyle、outlineLvl 与文本；级别归属由 TS 侧裁决。
 */
export interface RawParagraph {
  /** 全局物理指针（`<part>!<structuralPath>`）。 */
  pointer: string;
  part: string;
  /** 部件内结构路径，例如 `/w:document/w:body/w:p[3]`。 */
  path: string;
  /** 在同名兄弟中的序号（0 基）。 */
  ordinal: number;
  paraId: string | null;
  /** 归一化后的纯文本。 */
  text: string;
  styleId: string | null;
  styleName: string | null;
  /** 由 styles.xml 的继承链解析出的 outlineLvl（0..8），无则 null。 */
  outlineLevel: number | null;
  /** 由样式名推断出的级别（如 "Heading 2" -> 2），无则 null。 */
  styleNameLevel: number | null;
  /** 该段使用的样式是否属于标题样式族（引擎的原始判断，不含级别裁决）。 */
  headingStyle: boolean;
  list: RawListInfo | null;
  /** 本段内出现的批注引用 id。 */
  commentRefs: string[];
  /** 本段内出现的脚注/尾注引用 id。 */
  footnoteRefs: string[];
  endnoteRefs: string[];
}

/** 引擎观察到的原始单元格。 */
export interface RawCell {
  gridSpan: number;
  vMerge: 'restart' | 'continue' | null;
  paragraphs: RawParagraph[];
}

/** 引擎观察到的原始行。 */
export interface RawRow {
  cells: RawCell[];
}

/** 引擎观察到的原始表格。 */
export interface RawTable {
  pointer: string;
  part: string;
  path: string;
  ordinal: number;
  /** 表格自身的纯文本（各单元格以换行连接），仅用于指纹。 */
  text: string;
  /** tblGrid 声明的逻辑列数；无法判定时为 0。 */
  gridColumns: number;
  rows: RawRow[];
}

/** 引擎观察到的正文块。 */
export type RawBlock =
  | ({ kind: 'paragraph' } & RawParagraph)
  | ({ kind: 'table' } & RawTable);

/**
 * 引擎观察到的原始注释（批注 / 脚注 / 尾注）。
 *
 * `anchorPointers` 是它在正文里锚定到的块指针列表：可能为空（孤儿注释）、
 * 可能有一个（正常情况）、也可能有多个（同一批注跨越多个段落）。
 */
export interface RawAnnotation {
  kind: AnnotationKind;
  reference: string;
  text: string;
  author: string | null;
  date: string | null;
  pointer: string;
  part: string;
  path: string;
  ordinal: number;
  anchorPointers: string[];
}

/**
 * 引擎上报的非致命问题。
 *
 * 字段名是 `code`（引擎自有的问题码），它与本模块的 `DocxParseErrorCode`
 * 是两套东西：引擎问题码需要经 `mapper.ts` 映射后才成为模块错误码或告警码。
 */
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
  extension: string | null;
  sizeBytes: number;
  /** 文件字节的 SHA-256（小写十六进制）。 */
  sha256: string;
  /** 三态：true=已加密，false=未加密，null=无法判断。 */
  encrypted: boolean | null;
  documentKind: DocumentKind;
  documentVariant: DocumentVariant | null;
  mediaType: string | null;
  partCount: number;
  parts: PartRecord[];
  relationships: Array<{
    sourcePart: string;
    id: string;
    type: string;
    target: string;
    targetMode: 'Internal' | 'External';
  }>;
  /** 主文档部件名；无法判定时为空字符串。 */
  mainPart: string;
  /** 正文块（按文档顺序）。 */
  blocks: RawBlock[];
  /** 批注 / 脚注 / 尾注。 */
  annotations: RawAnnotation[];
  /** 命中的首个资源预算名；null 表示未超限。 */
  limitHit: string | null;
  /** 引擎上报的问题列表（非致命）。 */
  issues: ParseIssue[];
  /** 致命错误；非 null 时其余字段仅供参考。 */
  error: { kind: string; message: string } | null;
}

/* -------------------------------------------------------------------------- */
/* 边界校验（信任边界）                                                         */
/* -------------------------------------------------------------------------- */

/**
 * 当引擎输出的形状不符合 `ParseResult` 时抛出。
 *
 * 之所以单独定义而不复用 `DocxParseError`：它属于「协议层」错误，
 * 由 `adapter.ts` 捕获后翻译成 `ENGINE_PROTOCOL_ERROR`，
 * 从而让领域层不依赖模块的错误码定义（保持依赖方向单一）。
 */
export class ParseProtocolError extends Error {
  constructor(
    message: string,
    /** 出错位置，例如 `$.blocks[3].text`。 */
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

function asOptionalNumber(value: unknown, pointer: string): number | null {
  if (value === null || value === undefined) return null;
  return asNumber(value, pointer);
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

/** 校验枚举取值。 */
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

/** 校验可选枚举：允许 null / undefined。 */
function asOptionalEnum<T extends readonly string[]>(
  value: unknown,
  pointer: string,
  allowed: T,
): T[number] | null {
  if (value === null || value === undefined) return null;
  return asEnum(value, pointer, allowed);
}

/** 校验字符串数组（逐项断言类型）。 */
function asStringArray(value: unknown, pointer: string): string[] {
  return asArray(value, pointer).map((entry, index) => asString(entry, `${pointer}[${index}]`));
}

/** 解析列表编号信息。 */
function asListInfo(value: unknown, pointer: string): RawListInfo | null {
  if (value === null || value === undefined) return null;
  const record = asRecord(value, pointer);
  return {
    numId: asString(record.numId, `${pointer}.numId`),
    level: asNumber(record.level, `${pointer}.level`),
  };
}

/** 解析一个原始段落。 */
function asParagraph(value: unknown, pointer: string): RawParagraph {
  const record = asRecord(value, pointer);
  return {
    pointer: asString(record.pointer, `${pointer}.pointer`),
    part: asString(record.part, `${pointer}.part`),
    path: asString(record.path, `${pointer}.path`),
    ordinal: asNumber(record.ordinal, `${pointer}.ordinal`),
    paraId: asOptionalString(record.paraId, `${pointer}.paraId`),
    text: asString(record.text, `${pointer}.text`),
    styleId: asOptionalString(record.styleId, `${pointer}.styleId`),
    styleName: asOptionalString(record.styleName, `${pointer}.styleName`),
    outlineLevel: asOptionalNumber(record.outlineLevel, `${pointer}.outlineLevel`),
    styleNameLevel: asOptionalNumber(record.styleNameLevel, `${pointer}.styleNameLevel`),
    headingStyle: asBoolean(record.headingStyle, `${pointer}.headingStyle`),
    list: asListInfo(record.list, `${pointer}.list`),
    commentRefs: asStringArray(record.commentRefs, `${pointer}.commentRefs`),
    footnoteRefs: asStringArray(record.footnoteRefs, `${pointer}.footnoteRefs`),
    endnoteRefs: asStringArray(record.endnoteRefs, `${pointer}.endnoteRefs`),
  };
}

/** 解析一个正文块（段落或表格）。 */
function asBlock(value: unknown, pointer: string): RawBlock {
  const record = asRecord(value, pointer);
  const kind = asEnum(record.kind, `${pointer}.kind`, ['paragraph', 'table'] as const);

  if (kind === 'paragraph') {
    return { kind: 'paragraph', ...asParagraph(record, pointer) };
  }

  const rows = asArray(record.rows, `${pointer}.rows`).map((rowEntry, rowIndex) => {
    const rowPointer = `${pointer}.rows[${rowIndex}]`;
    const row = asRecord(rowEntry, rowPointer);
    const cells = asArray(row.cells, `${rowPointer}.cells`).map((cellEntry, cellIndex) => {
      const cellPointer = `${rowPointer}.cells[${cellIndex}]`;
      const cell = asRecord(cellEntry, cellPointer);
      return {
        gridSpan: asNumber(cell.gridSpan, `${cellPointer}.gridSpan`),
        vMerge: asOptionalEnum(cell.vMerge, `${cellPointer}.vMerge`, [
          'restart',
          'continue',
        ] as const),
        paragraphs: asArray(cell.paragraphs, `${cellPointer}.paragraphs`).map((pEntry, pIndex) =>
          asParagraph(pEntry, `${cellPointer}.paragraphs[${pIndex}]`),
        ),
      };
    });
    return { cells };
  });

  return {
    kind: 'table',
    pointer: asString(record.pointer, `${pointer}.pointer`),
    part: asString(record.part, `${pointer}.part`),
    path: asString(record.path, `${pointer}.path`),
    ordinal: asNumber(record.ordinal, `${pointer}.ordinal`),
    text: asString(record.text, `${pointer}.text`),
    gridColumns: asNumber(record.gridColumns, `${pointer}.gridColumns`),
    rows,
  };
}

/** 解析一条原始注释。 */
function asAnnotation(value: unknown, pointer: string): RawAnnotation {
  const record = asRecord(value, pointer);
  return {
    kind: asEnum(record.kind, `${pointer}.kind`, ANNOTATION_KINDS),
    reference: asString(record.reference, `${pointer}.reference`),
    text: asString(record.text, `${pointer}.text`),
    author: asOptionalString(record.author, `${pointer}.author`),
    date: asOptionalString(record.date, `${pointer}.date`),
    pointer: asString(record.pointer, `${pointer}.pointer`),
    part: asString(record.part, `${pointer}.part`),
    path: asString(record.path, `${pointer}.path`),
    ordinal: asNumber(record.ordinal, `${pointer}.ordinal`),
    anchorPointers: asStringArray(record.anchorPointers, `${pointer}.anchorPointers`),
  };
}

/**
 * 把原始引擎载荷校验并归一化为 `ParseResult`。
 *
 * 这是整个模块的信任边界：载荷来自子进程，被视为不可信输入，
 * 每个字段在使用前都被检查，因此畸形输出只能变成受控的协议错误。
 *
 * @throws ParseProtocolError 当载荷形状不符合协议时
 */
export function parseParseResult(payload: unknown): ParseResult {
  const root = asRecord(payload, '$');

  const parts = asArray(root.parts, '$.parts').map((entry, index) => {
    const record = asRecord(entry, `$.parts[${index}]`);
    return {
      name: asString(record.name, `$.parts[${index}].name`),
      sizeBytes: asNumber(record.sizeBytes, `$.parts[${index}].sizeBytes`),
      compressedSize: asNumber(record.compressedSize, `$.parts[${index}].compressedSize`),
    };
  });

  const relationships = asArray(root.relationships, '$.relationships').map((entry, index) => {
    const record = asRecord(entry, `$.relationships[${index}]`);
    return {
      sourcePart: asString(record.sourcePart, `$.relationships[${index}].sourcePart`),
      id: asString(record.id, `$.relationships[${index}].id`),
      type: asString(record.type, `$.relationships[${index}].type`),
      target: asString(record.target, `$.relationships[${index}].target`),
      targetMode: asEnum(record.targetMode, `$.relationships[${index}].targetMode`, [
        'Internal',
        'External',
      ] as const),
    };
  });

  const blocks = asArray(root.blocks, '$.blocks').map((entry, index) =>
    asBlock(entry, `$.blocks[${index}]`),
  );

  const annotations = asArray(root.annotations, '$.annotations').map((entry, index) =>
    asAnnotation(entry, `$.annotations[${index}]`),
  );

  const issues = asArray(root.issues, '$.issues').map((entry, index) => {
    const record = asRecord(entry, `$.issues[${index}]`);
    return {
      code: asString(record.code, `$.issues[${index}].code`),
      message: asString(record.message, `$.issues[${index}].message`),
      // path 为可选字段：仅在存在时才写入，避免出现 `path: undefined`。
      ...(record.path === undefined || record.path === null
        ? {}
        : { path: asString(record.path, `$.issues[${index}].path`) }),
    };
  });

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
    partCount: asNumber(root.partCount, '$.partCount'),
    parts,
    relationships,
    mainPart: asString(root.mainPart, '$.mainPart'),
    blocks,
    annotations,
    limitHit: asOptionalString(root.limitHit, '$.limitHit'),
    issues,
    error,
  };
}

/* -------------------------------------------------------------------------- */
/* 领域辅助函数（仅做汇总，不做风险判定）                                       */
/* -------------------------------------------------------------------------- */

/** 统计表格数量。 */
export function countTables(result: ParseResult): number {
  return result.blocks.filter((block) => block.kind === 'table').length;
}

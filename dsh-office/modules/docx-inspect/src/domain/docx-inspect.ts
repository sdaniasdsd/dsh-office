/**
 * docx-inspect 的领域模型。
 *
 * `ProbeResult` 是对产物的「与引擎无关」的描述：Python 桥接把它序列化成 JSON 输出，
 * 除此之外没有任何东西跨越进程边界。容器、宏、外部引用、嵌入对象都在这里建模；
 * 映射到 office-core 的 IR 发生在 `mapper.ts`，这样领域形状就能独立于 IR 版本演进。
 *
 * 分层意图：
 *   engine(进程外) → probeResult(本文件，中立) → mapper(归一化) → office-core IR
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
export const DOCUMENT_VARIANTS = ['document', 'template', 'macroEnabled', 'macroEnabledTemplate'] as const;
export type DocumentVariant = (typeof DOCUMENT_VARIANTS)[number];

/** 外部引用的分类。它决定 `verify` 时评估哪一条策略项。 */
export const EXTERNAL_CATEGORIES = [
  'attachedTemplate',
  'hyperlink',
  'externalImage',
  'externalObject',
  'dataSource',
  'dde',
  'other',
] as const;
export type ExternalCategory = (typeof EXTERNAL_CATEGORIES)[number];

/** 嵌入对象的种类。 */
export const EMBEDDED_KINDS = ['oleEmbedding', 'package', 'altChunk', 'activeX', 'flash'] as const;
export type EmbeddedKind = (typeof EMBEDDED_KINDS)[number];

/** 包内部件记录。 */
export interface PartRecord {
  name: string;
  sizeBytes: number;
  compressedSize: number;
}

/** 宏指标（仅「存在性」，不含宏内容）。 */
export interface MacroIndicator {
  kind: 'vba' | 'xlm';
  part: string;
  sizeBytes: number;
}

/**
 * 宏代码的语义分析结论（可选增强能力，依赖 Python 侧的 oletools）。
 *
 * 它是「存在性指标」的升级：`MacroIndicator` 只说明「有宏」，
 * 而本结构说明「宏在干什么」。判定沿用 oletools 的 MacroRaptor 三要素：
 *
 *     suspicious = autoExec ∧ (write ∨ execute)
 *
 * 之所以不做更深的语义分析：真正的恶意宏几乎必然命中 A 与 X/W 的组合，
 * 而这三要素均可解释、可审计，误报率也远低于「看到宏就报警」。
 */
export interface MacroAnalysis {
  /** VBA 模块名（取自 olevba 解析出的 VBA 内部文件名）。 */
  module: string;
  /** A：存在自动执行触发点（AutoOpen / Document_Open / Workbook_Open 等）。 */
  autoExec: boolean;
  /** W：存在写文件或写内存行为（如 CreateTextFile / VirtualAlloc）。 */
  write: boolean;
  /** X：存在执行外部程序或命令的行为（如 Shell / CreateObject）。 */
  execute: boolean;
  /** 最终判定：A ∧ (W ∨ X)。 */
  suspicious: boolean;
  /** 形如 "A-X" 的三字符标志（A/W/X，未命中为 "-"），便于人读与检索。 */
  flags: string;
  /** 命中的关键字列表，作为判定证据回传，便于上层审计与展示。 */
  matches: string[];
}

/** 外部引用指标。 */
export interface ExternalReference {
  sourcePart: string;
  relationshipId: string;
  relationshipType: string;
  target: string;
  category: ExternalCategory;
}

/** 嵌入对象指标。 */
export interface EmbeddedObject {
  part: string;
  kind: EmbeddedKind;
  mediaType: string | null;
  sizeBytes: number;
  name: string | null;
}

/**
 * 引擎上报的非致命问题。
 *
 * 注意字段名是 `code`（引擎自有的问题码），它与本模块的 `DocxInspectErrorCode`
 * 是两套东西：引擎问题码需要经 `mapper.ts` 映射后才成为模块错误码或告警码。
 */
export interface ProbeIssue {
  code: string;
  message: string;
  path?: string;
}

/**
 * 探测结果：引擎与模块之间的唯一数据契约。
 *
 * 所有字段都必须可 JSON 序列化（这是跨进程传输的前提）。
 */
export interface ProbeResult {
  /** 探测协议版本，便于引擎与模块独立演进时做兼容判断。 */
  probeVersion: number;
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
  partCount: number;
  parts: PartRecord[];
  relationships: Array<{
    sourcePart: string;
    id: string;
    type: string;
    target: string;
    targetMode: 'Internal' | 'External';
  }>;
  macroIndicators: MacroIndicator[];
  /** 宏代码语义分析结果（可选增强；未启用、依赖缺失或无宏时为空数组）。 */
  macroAnalyses: MacroAnalysis[];
  externalReferences: ExternalReference[];
  embeddedObjects: EmbeddedObject[];
  /** 命中的 DDE/DDEAUTO 字段指令（截断后的文本）。 */
  ddeFields: string[];
  /** 首个被触发的资源预算名；null 表示未超限。 */
  limitHit: string | null;
  /** 引擎上报的问题列表（非致命）。 */
  issues: ProbeIssue[];
  /** 致命错误；非 null 时其余字段仅供参考。 */
  error: { kind: string; message: string } | null;
}

/* -------------------------------------------------------------------------- */
/* 边界校验                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * 当引擎输出的形状不符合 `ProbeResult` 时抛出。
 *
 * 之所以单独定义而不复用 `DocxInspectError`：
 * 它属于「协议层」错误，由 `adapter.ts` 捕获后翻译成 `ENGINE_PROTOCOL_ERROR`，
 * 从而让领域层不依赖模块的错误码定义（保持依赖方向单一）。
 */
export class ProbeProtocolError extends Error {
  constructor(
    message: string,
    /** 出错位置，例如 `$.parts[3].name`。 */
    readonly pointer: string,
  ) {
    super(`${message} (at ${pointer})`);
    this.name = 'ProbeProtocolError';
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
    throw new ProbeProtocolError('expected object', pointer);
  }
  return value as Record<string, unknown>;
}

function asArray(value: unknown, pointer: string): unknown[] {
  if (!Array.isArray(value)) {
    throw new ProbeProtocolError('expected array', pointer);
  }
  return value;
}

function asString(value: unknown, pointer: string): string {
  if (typeof value !== 'string') {
    throw new ProbeProtocolError('expected string', pointer);
  }
  return value;
}

function asOptionalString(value: unknown, pointer: string): string | null {
  if (value === null || value === undefined) return null;
  return asString(value, pointer);
}

function asNumber(value: unknown, pointer: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new ProbeProtocolError('expected finite number', pointer);
  }
  return value;
}

/** 严格布尔断言：只接受真正的 boolean，不接受 0/1 或字符串。 */
function asBoolean(value: unknown, pointer: string): boolean {
  if (typeof value !== 'boolean') {
    throw new ProbeProtocolError('expected boolean', pointer);
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
    throw new ProbeProtocolError(`unexpected value "${raw}"`, pointer);
  }
  return raw as T[number];
}

/**
 * 把原始引擎载荷校验并归一化为 `ProbeResult`。
 *
 * 这是整个模块的信任边界：载荷来自子进程，被视为不可信输入，
 * 每个字段在使用前都被检查，因此畸形输出只能变成受控的协议错误，
 * 而不会污染后续逻辑。
 *
 * @throws ProbeProtocolError 当载荷形状不符合协议时
 */
export function parseProbeResult(payload: unknown): ProbeResult {
  const root = asRecord(payload, '$');

  // --- 包内部件列表 ------------------------------------------------------ //
  const parts = asArray(root.parts, '$.parts').map((entry, index) => {
    const record = asRecord(entry, `$.parts[${index}]`);
    return {
      name: asString(record.name, `$.parts[${index}].name`),
      sizeBytes: asNumber(record.sizeBytes, `$.parts[${index}].sizeBytes`),
      compressedSize: asNumber(record.compressedSize, `$.parts[${index}].compressedSize`),
    };
  });

  // --- OPC 关系列表 ------------------------------------------------------ //
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

  // --- 宏指标 ------------------------------------------------------------ //
  const macroIndicators = asArray(root.macroIndicators, '$.macroIndicators').map((entry, index) => {
    const record = asRecord(entry, `$.macroIndicators[${index}]`);
    return {
      kind: asEnum(record.kind, `$.macroIndicators[${index}].kind`, ['vba', 'xlm'] as const),
      part: asString(record.part, `$.macroIndicators[${index}].part`),
      sizeBytes: asNumber(record.sizeBytes, `$.macroIndicators[${index}].sizeBytes`),
    };
  });

  // --- 宏语义分析（可选增强） -------------------------------------------- //
  // 逐字段校验：这部分数据最终来自第三方库（oletools），更需要按不可信输入对待。
  const macroAnalyses = asArray(root.macroAnalyses, '$.macroAnalyses').map((entry, index) => {
    const record = asRecord(entry, `$.macroAnalyses[${index}]`);
    return {
      module: asString(record.module, `$.macroAnalyses[${index}].module`),
      autoExec: asBoolean(record.autoExec, `$.macroAnalyses[${index}].autoExec`),
      write: asBoolean(record.write, `$.macroAnalyses[${index}].write`),
      execute: asBoolean(record.execute, `$.macroAnalyses[${index}].execute`),
      suspicious: asBoolean(record.suspicious, `$.macroAnalyses[${index}].suspicious`),
      flags: asString(record.flags, `$.macroAnalyses[${index}].flags`),
      matches: asArray(record.matches, `$.macroAnalyses[${index}].matches`).map((item, at) =>
        asString(item, `$.macroAnalyses[${index}].matches[${at}]`),
      ),
    };
  });

  // --- 外部引用 ---------------------------------------------------------- //
  const externalReferences = asArray(root.externalReferences, '$.externalReferences').map(
    (entry, index) => {
      const record = asRecord(entry, `$.externalReferences[${index}]`);
      return {
        sourcePart: asString(record.sourcePart, `$.externalReferences[${index}].sourcePart`),
        relationshipId: asString(
          record.relationshipId,
          `$.externalReferences[${index}].relationshipId`,
        ),
        relationshipType: asString(
          record.relationshipType,
          `$.externalReferences[${index}].relationshipType`,
        ),
        target: asString(record.target, `$.externalReferences[${index}].target`),
        category: asEnum(
          record.category,
          `$.externalReferences[${index}].category`,
          EXTERNAL_CATEGORIES,
        ),
      };
    },
  );

  // --- 嵌入对象 ---------------------------------------------------------- //
  const embeddedObjects = asArray(root.embeddedObjects, '$.embeddedObjects').map((entry, index) => {
    const record = asRecord(entry, `$.embeddedObjects[${index}]`);
    return {
      part: asString(record.part, `$.embeddedObjects[${index}].part`),
      kind: asEnum(record.kind, `$.embeddedObjects[${index}].kind`, EMBEDDED_KINDS),
      mediaType: asOptionalString(record.mediaType, `$.embeddedObjects[${index}].mediaType`),
      sizeBytes: asNumber(record.sizeBytes, `$.embeddedObjects[${index}].sizeBytes`),
      name: asOptionalString(record.name, `$.embeddedObjects[${index}].name`),
    };
  });

  // --- 问题列表 ---------------------------------------------------------- //
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

  // --- 致命错误 ---------------------------------------------------------- //
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
    probeVersion: asNumber(root.probeVersion, '$.probeVersion'),
    xmlBackend: asString(root.xmlBackend, '$.xmlBackend'),
    container: asEnum(root.container, '$.container', CONTAINER_KINDS),
    extension: asOptionalString(root.extension, '$.extension'),
    sizeBytes: asNumber(root.sizeBytes, '$.sizeBytes'),
    sha256: asString(root.sha256, '$.sha256'),
    encrypted: asOptionalBoolean(root.encrypted, '$.encrypted'),
    documentKind: asEnum(root.documentKind, '$.documentKind', DOCUMENT_KINDS),
    documentVariant:
      root.documentVariant === null || root.documentVariant === undefined
        ? null
        : asEnum(root.documentVariant, '$.documentVariant', DOCUMENT_VARIANTS),
    mediaType: asOptionalString(root.mediaType, '$.mediaType'),
    partCount: asNumber(root.partCount, '$.partCount'),
    parts,
    relationships,
    macroIndicators,
    macroAnalyses,
    externalReferences,
    embeddedObjects: dedupeEmbedded(embeddedObjects),
    ddeFields: asArray(root.ddeFields, '$.ddeFields').map((entry, index) =>
      asString(entry, `$.ddeFields[${index}]`),
    ),
    limitHit: asOptionalString(root.limitHit, '$.limitHit'),
    issues,
    error,
  };
}

/**
 * 去重嵌入对象。
 *
 * 引擎可能从多条发现路径（目录前缀、Content-Type、文件后缀）命中同一个部件，
 * 这里按「种类 + 部件」去重，保证 `mapper` 收到的清单唯一且稳定。
 */
function dedupeEmbedded(objects: EmbeddedObject[]): EmbeddedObject[] {
  const seen = new Set<string>();
  const result: EmbeddedObject[] = [];
  for (const object of objects) {
    // 使用 \u0000 作为分隔符：它不可能出现在 OOXML 部件名中，避免拼接歧义。
    const key = `${object.kind}\u0000${object.part}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(object);
  }
  return result;
}

/* -------------------------------------------------------------------------- */
/* 领域辅助函数（仅做指标汇总，不做风险判定）                                   */
/* -------------------------------------------------------------------------- */

/** 是否携带宏（VBA 或 XLM）。 */
export function hasMacros(probe: ProbeResult): boolean {
  return probe.macroIndicators.length > 0;
}

/** 是否存在外部引用。 */
export function hasExternalReferences(probe: ProbeResult): boolean {
  return probe.externalReferences.length > 0;
}

/**
 * 是否存在嵌入对象。
 *
 * 注意排除 `altChunk`：它有独立的能力开关与告警码，
 * 混在一起会让「嵌入对象」这个语义变得含糊。
 */
export function hasEmbeddedObjects(probe: ProbeResult): boolean {
  return probe.embeddedObjects.some((object) => object.kind !== 'altChunk');
}

/** 是否包含 altChunk（可引入外部内容块）。 */
export function hasAltChunks(probe: ProbeResult): boolean {
  return probe.embeddedObjects.some((object) => object.kind === 'altChunk');
}

/** 是否包含 ActiveX 控件。 */
export function hasActiveX(probe: ProbeResult): boolean {
  return probe.embeddedObjects.some((object) => object.kind === 'activeX');
}

/** 是否命中 DDE/DDEAUTO 字段。 */
export function hasDdeFields(probe: ProbeResult): boolean {
  return probe.ddeFields.length > 0;
}

/** 按分类筛选外部引用。 */
export function externalTargetsByCategory(
  probe: ProbeResult,
  category: ExternalCategory,
): ExternalReference[] {
  return probe.externalReferences.filter((reference) => reference.category === category);
}

/**
 * 对一个外部关系目标做分类。
 *
 * 这是纯字符串/类型判断：docx-inspect【只做标注】，既不抓取目标，
 * 也不判断它是否可接受（后者属于 office-safety 的职责）。
 *
 * 注意：这里的判断逻辑必须与 Python 引擎侧的 `classify_external` 保持一致，
 * 否则同一份文档在不同链路上会得到不同分类。
 */
export function classifyExternalTarget(relationshipType: string, target: string): ExternalCategory {
  const type = relationshipType.toLowerCase();
  if (type.endsWith('/attachedtemplate')) return 'attachedTemplate';
  if (type.endsWith('/hyperlink')) return 'hyperlink';
  if (type.endsWith('/image')) return 'externalImage';
  if (type.endsWith('/oleobject') || type.endsWith('/package')) return 'externalObject';
  if (type.includes('mailmerge') || type.endsWith('/datasource')) return 'dataSource';
  if (type.endsWith('/framefile')) return 'externalObject';
  if (/^dde/i.test(target) || /ddeauto/i.test(target)) return 'dde';
  return 'other';
}

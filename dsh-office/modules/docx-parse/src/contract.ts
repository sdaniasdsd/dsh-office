/**
 * docx-parse 公开契约 —— 【已冻结的表面】。
 *
 * 本文件导出的所有内容，都属于本模块在 Profile 中的「注册表面」。修改它们对下游
 * 消费者而言是破坏性变更，因此必须把它当作带版本号的接口来对待。
 *
 * 本模块的核心职责是产出【双 IR】：同一份 DOCX 被解析成两个互相印证、且能通过
 * 一张对应哈希表精确互指的视图：
 *
 *   1. 【语义视图】(SemanticDocument) —— 给人和 LLM 看的「地图」：
 *      标题层级、段落、表格行列、批注、脚注，都是归一化后的意义单位。
 *   2. 【物理视图】(PhysicalDocument) —— 给写入端看的「施工图纸」：
 *      OOXML 里真实存在的节点及其结构路径与内容指纹。
 *   3. 【对应表】(SourceMap) —— 双视图之间的「账本」：
 *      语义节点 ↔ 物理节点的映射，并附带多键索引（id / 指针 / 内容指纹）。
 *
 * 模块强制保证的三条不变量（由 `verifier.ts` 兜底检查）：
 *   1. 任何「输入」都可 JSON 序列化；
 *   2. 任何「输出」都可 JSON 序列化，且不包含任何底层引擎对象
 *      （不允许出现 Python 句柄、`zipfile`/`lxml` 实例、Buffer、类实例）；
 *   3. 错误分类、artifact 引用、验证报告与【身份方案】由本模块自己拥有，
 *      底层库无法直接改写它们。引擎只能上报「观察到的原始事实」，
 *      它无权决定语义/物理视图的形状，也无权决定节点 id。
 *
 * 上游模块仅以 `import type` 方式消费，因此本模块与 office-core / office-safety
 * 之间不存在运行时耦合。
 */
import type { ArtifactRef, FormatIR, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { VerificationReportLike } from 'office-test-kit';

/* -------------------------------------------------------------------------- */
/* JSON 原语                                                                   */
/* -------------------------------------------------------------------------- */

/** JSON 标量。 */
export type JsonPrimitive = string | number | boolean | null;
/** 任意可 JSON 序列化的值（递归定义）。 */
export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };
/** 可 JSON 序列化的普通对象。 */
export type JsonObject = { [key: string]: JsonValue };

/* -------------------------------------------------------------------------- */
/* 身份标识                                                                    */
/* -------------------------------------------------------------------------- */

/** 模块在 Profile 中的唯一 id。 */
export const DOCX_PARSE_MODULE_ID = 'docx-parse' as const;
export type DocxParseModuleId = typeof DOCX_PARSE_MODULE_ID;

/** 本模块对外发布的三个接口（见 spec 第三节）。 */
export const DOCX_PARSE_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export type DocxParseCapability = (typeof DOCX_PARSE_CAPABILITIES)[number];

/** 操作名与能力名同构（让输入里的 operation 字段在类型上自解释）。 */
export type DocxOperation = DocxParseCapability;

/* -------------------------------------------------------------------------- */
/* 错误码                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 稳定的错误分类体系。
 *
 * 关键设计：底层引擎只能上报一个「问题（issue）」，而「问题 → 错误码」的映射
 * 由 `mapper.ts` 决定，引擎无权选择错误码。这样即使更换引擎实现，
 * 上层看到的错误语义也完全一致。
 */
export const DOCX_PARSE_ERROR_CODES = {
  /** 调用方传参非法：字段缺失/类型错误/不可序列化。 */
  INVALID_INPUT: 'INVALID_INPUT',
  /** 引用的 artifact 不存在或不可读。 */
  ARTIFACT_NOT_FOUND: 'ARTIFACT_NOT_FOUND',
  /** 引用指向默认解析器够不到的位置（http、memory 等）。 */
  UNSUPPORTED_ARTIFACT_URI: 'UNSUPPORTED_ARTIFACT_URI',
  /** 声明的扩展名/MIME 与字节内容不符，或包并非 WordprocessingML。 */
  FORMAT_MISMATCH: 'FORMAT_MISMATCH',
  /** 容器可识别但超出本模块范围（OLE/CFB、RTF、加密 OOXML 等）。 */
  UNSUPPORTED_CONTAINER: 'UNSUPPORTED_CONTAINER',
  /** 安全策略拒绝了该 artifact（`inspect` 快速失败）。 */
  SAFETY_POLICY_DENIED: 'SAFETY_POLICY_DENIED',
  /** 无法启动 Python 解释器或解析脚本。 */
  ENGINE_UNAVAILABLE: 'ENGINE_UNAVAILABLE',
  /** 引擎进程以非零码退出。 */
  ENGINE_FAILED: 'ENGINE_FAILED',
  /** 引擎超过 `config.timeoutMs`。 */
  ENGINE_TIMEOUT: 'ENGINE_TIMEOUT',
  /** 引擎输出不符合解析协议（JSON 非法或结构不符）。 */
  ENGINE_PROTOCOL_ERROR: 'ENGINE_PROTOCOL_ERROR',
  /** 包被截断或结构损坏，无法解析。 */
  PARSE_FAILED: 'PARSE_FAILED',
  /** 超出 `limits` 配置的预算（解压炸弹防护）。 */
  LIMIT_EXCEEDED: 'LIMIT_EXCEEDED',
  /**
   * 双 IR 的对应关系【不成立】：语义节点找不到物理锚点，或反之。
   *
   * 这是本模块独有的一类失败——它不是「文档坏了」，而是「两份视图对不上号」。
   * 一旦出现，说明双 IR 的可靠性被破坏，宁可显式失败也不能返回一份
   * 看起来完整、实则无法定位的 IR。
   */
  SOURCEMAP_INCOMPLETE: 'SOURCEMAP_INCOMPLETE',
  /** `verify` 无法产出报告，或某项强制检查硬失败。 */
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;

/** 错误码联合类型。 */
export type DocxParseErrorCode = keyof typeof DOCX_PARSE_ERROR_CODES;

/* -------------------------------------------------------------------------- */
/* 告警码                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 告警码表。
 *
 * 注意区分：错误码表示「调用失败」，告警码表示「调用成功但结果需要被注意」。
 * 解析过程中遇到本模块尚未建模的 OOXML 元素时，一律降级为告警而不是失败——
 * 这保证了「部分解析」仍然可用，同时把「哪些内容没被理解」如实告知上层。
 */
export const DOCX_PARSE_WARNING_CODES = {
  /** 文档结构合法，但正文中没有任何可见块。 */
  EMPTY_DOCUMENT: 'EMPTY_DOCUMENT',
  /** 只解析了部分内容（存在跳过或未建模的元素）。 */
  PARTIALLY_PARSED: 'PARTIALLY_PARSED',
  /** 遇到本模块未建模的 OOXML 元素。 */
  UNSUPPORTED_CONTENT: 'UNSUPPORTED_CONTENT',
  /** 命中资源预算但被宽免（`enforceLimits` 关闭）。 */
  LIMIT_APPLIED: 'LIMIT_APPLIED',
  /** 样式表缺失（styles.xml 不可读），标题层级退化为纯推断。 */
  STYLE_FALLBACK: 'STYLE_FALLBACK',
  /** 标题层级不是由 outlineLvl 明确给出，而是推断所得。 */
  HEADING_LEVEL_INFERRED: 'HEADING_LEVEL_INFERRED',
  /** 某个锚点只有一个选择器可用（冗余不足，抗漂移能力弱）。 */
  ANCHOR_INCOMPLETE: 'ANCHOR_INCOMPLETE',
  /** 批注在正文中找不到锚点，或锚点指向不存在的块。 */
  COMMENT_ORPHANED: 'COMMENT_ORPHANED',
  /** 脚注/尾注在正文中找不到引用点。 */
  NOTE_ORPHANED: 'NOTE_ORPHANED',
  /** 产物已加密，无法解析内容。 */
  ENCRYPTED_ARTIFACT: 'ENCRYPTED_ARTIFACT',
  /** 验证未通过（作为告警广播，便于只消费 warnings 的上层感知）。 */
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;

/** 告警码联合类型。 */
export type DocxParseWarningCode = keyof typeof DOCX_PARSE_WARNING_CODES;

/* -------------------------------------------------------------------------- */
/* 配置（由 Profile 注入；模块从不读取全局配置）                                */
/* -------------------------------------------------------------------------- */

/** 引擎相关配置。 */
export interface EngineConfig {
  /** 引擎族。目前只实现了 Python 桥接。 */
  driver: 'python';
  /** 启动解析脚本所用的解释器（可为命令名或绝对路径）。 */
  pythonPath: string;
  /** 解析脚本的绝对路径；省略时取 adapter.ts 同目录下的 docx_parse.py。 */
  scriptPath?: string;
  /** 传给子进程的额外环境变量。 */
  env?: Record<string, string>;
}

/** 资源预算，用于抵御 zip 炸弹与超大 OOXML 包。 */
export interface LimitConfig {
  /** 检查的归档条目数上限，超出即中止。 */
  maxArchiveEntries: number;
  /** 单个条目解压后的大小上限。 */
  maxEntryUncompressedBytes: number;
  /** 所有条目解压后的总大小上限。 */
  maxTotalUncompressedBytes: number;
  /** 解析的 `.rels` 关系条数上限。 */
  maxRelationships: number;
  /** 正文块（段落 + 表格）数量上限。 */
  maxBlocks: number;
  /** 单表单元格数量上限。 */
  maxTableCells: number;
  /** 解析的批注数量上限。 */
  maxAnnotations: number;
}

/**
 * 能力开关。
 *
 * 关闭某项后，对应的视图部分【不会被解析】，因此其锚点与统计也会缺席。
 * 这一点很重要：验证器必须能区分「没有该项」与「没有解析该项」，
 * 否则会产生虚假的安全感（详见 `verifier.ts` 的 skip 语义）。
 */
export interface FeatureFlags {
  /** 解析标题层级（需要读取 styles.xml 的 outlineLvl 继承链）。 */
  parseHeadings: boolean;
  /** 解析表格结构（行列、横向/纵向合并、嵌套段落）。 */
  parseTables: boolean;
  /** 解析样式引用（pStyle 的 id 与可读名）。 */
  parseStyles: boolean;
  /** 解析批注（comments.xml 及正文锚点）。 */
  parseComments: boolean;
  /** 解析脚注与尾注（footnotes.xml / endnotes.xml）。 */
  parseFootnotes: boolean;
  /**
   * 是否构建身份锚点与对应哈希表。
   *
   * 关闭后 `execute` 只返回「视图内容」，`sourceMap.entries` 为空，
   * 且 `verify` 的对应关系检查会判为 skip。这是给「纯读取、不需要回写定位」
   * 的消费方准备的省事开关。
   */
  resolveAnchors: boolean;
  /** 为 false 时，引擎只上报超限但不中止（用于「尽力而为」的宽限场景）。 */
  enforceLimits: boolean;
  /**
   * 是否额外收集**表现层观察**：段落对齐与逐 run 字符格式。
   *
   * 默认关闭，因为语义视图刻意只承载「这一段说的是什么、它是谁」；字体、字号、
   * 加粗与对齐回答的是「它长什么样」，而 run 级数据会显著放大 IR。需要按参照文档
   * 重建版面的消费方（如参照生成）打开它，其它消费方不为此付费。
   */
  parseFormatting: boolean;
}

/** 模块完整配置：由 Profile 注入，模块不读全局配置。 */
export interface ModuleConfig {
  engine: EngineConfig;
  limits: LimitConfig;
  timeoutMs: number;
  featureFlags: FeatureFlags;
}

/* -------------------------------------------------------------------------- */
/* 双 IR（本模块的核心交付物）                                                  */
/* -------------------------------------------------------------------------- */

/** 身份方案标识。改动锚点构成即意味着改身份，必须同时改动这个版本号。 */
export const SOURCE_MAP_SCHEME = 'composite-anchor-v1' as const;
export type SourceMapScheme = typeof SOURCE_MAP_SCHEME;

/** 双 IR 的视图数量常量。 */
export const DUAL_VIEW_COUNT = 2 as const;

/** 语义节点的稳定标识（十六进制摘要）。 */
export type SemanticId = string;

/**
 * 节点锚点 —— 双 IR 的「身份」载体。
 *
 * 设计要点：不赌任何单一选择器，而是【冗余取证】。
 * 三种选择器各自会在不同场景下失效，但组合起来能互相兜底：
 *
 *   | 选择器          | 作用           | 失效场景       | 由谁兜底        |
 *   | --------------- | -------------- | -------------- | --------------- |
 *   | paraId          | 原生稳定 id    | 可选、常缺失   | quote / path    |
 *   | quote           | 内容寻址       | 内容被编辑     | paraId / path   |
 *   | structuralPath  | 位置定位       | 增删导致漂移   | paraId / quote  |
 *
 * 这也是为什么 `ANCHOR_INCOMPLETE` 是一条告警：当某个节点只剩一个选择器可用时，
 * 它的抗漂移能力已经退化，上层有权知道。
 */
export interface NodeAnchor {
  /** 节点的 OOXML 本地名（p / tbl / comment 等）。 */
  kind: string;
  /** 承载该节点的部件名，例如 `word/document.xml`。 */
  part: string;
  /** 部件内的结构路径，例如 `/w:document/w:body/w:p[3]`。 */
  structuralPath: string;
  /** 在同名兄弟中的序号（0 基）。 */
  ordinal: number;
  /** OOXML 原生的 w14:paraId（存在时填写，否则 null）。 */
  paraId: string | null;
  /** 归一化文本的截断副本，作为内容选择器。 */
  quote: string;
  /** 内容指纹：sha256(kind ∥ 归一化文本)，用于内容寻址。 */
  digest: string;
}

/** 标题层级的来源，说明「这个级别是怎么来的」。 */
export type HeadingLevelSource = 'outlineLevel' | 'styleName' | 'inferred';

/** 段落的列表编号信息。 */
export interface SemanticListInfo {
  /** numPr 指向的 numId。 */
  numId: string;
  /** 层级（0 基）。 */
  level: number;
}

/** 语义块共有的字段。 */
export interface SemanticBlockBase {
  /** 稳定标识，可通过 SourceMap 反查物理节点。 */
  id: SemanticId;
  kind: string;
  anchor: NodeAnchor;
}

/** 语义视图中的标题。 */
export interface SemanticHeading extends SemanticBlockBase {
  kind: 'heading';
  /** 归一化后的标题级别，1..9。 */
  level: number;
  /** 级别来源：outlineLvl 明确给出 / 由样式名推断 / 纯启发式推断。 */
  levelSource: HeadingLevelSource;
  text: string;
  styleId: string | null;
  styleName: string | null;
}

/** 语义视图中的普通段落。 */
export interface SemanticParagraph extends SemanticBlockBase {
  kind: 'paragraph';
  text: string;
  styleId: string | null;
  styleName: string | null;
  /** 列表编号信息；非列表段落为 null。 */
  list: SemanticListInfo | null;
}

/** 表格单元格。 */
export interface SemanticCell {
  /** 单元格内纯文本（各段落以换行连接）。 */
  text: string;
  /** 横向合并跨度（gridSpan），默认 1。 */
  gridSpan: number;
  /** 纵向合并状态：restart=起始格，continue=被合并格，null=无纵向合并。 */
  vMerge: 'restart' | 'continue' | null;
  /** 单元格内的段落（含各自锚点，因此可以精确定位到单元格内部）。 */
  paragraphs: SemanticParagraph[];
}

/** 表格行。 */
export interface SemanticRow {
  cells: SemanticCell[];
}

/** 语义视图中的表格。 */
export interface SemanticTable extends SemanticBlockBase {
  kind: 'table';
  /** tblGrid 声明的逻辑列数。 */
  gridColumns: number;
  rows: SemanticRow[];
}

/** 语义块联合类型。 */
export type SemanticBlock = SemanticHeading | SemanticParagraph | SemanticTable;

/** 批注 / 脚注 / 尾注的种类。 */
export type SemanticAnnotationKind = 'comment' | 'footnote' | 'endnote';

/** 语义视图中的批注或注释。 */
export interface SemanticAnnotation {
  /** 稳定标识。 */
  id: SemanticId;
  kind: SemanticAnnotationKind;
  /** OOXML 原生引用号（comment w:id / footnote w:id）。 */
  reference: string;
  text: string;
  author: string | null;
  date: string | null;
  /** 锚定到哪个正文块；找不到对应块时为 null（并产生占位告警）。 */
  anchorBlockId: SemanticId | null;
  /** 该注释自身的锚点（位于 comments.xml / footnotes.xml 部件内）。 */
  anchor: NodeAnchor;
}

/** 语义视图：给人和 LLM 看的「地图」。 */
export interface SemanticDocument {
  view: 'semantic';
  blocks: SemanticBlock[];
  annotations: SemanticAnnotation[];
}

/** 物理视图中的单个节点：OOXML 里真实存在的元素。 */
export interface PhysicalNode {
  /** 全局唯一的物理指针，格式 `<part>!<structuralPath>`。 */
  pointer: string;
  part: string;
  /** 元素本地名。 */
  kind: string;
  ordinal: number;
  paraId: string | null;
  /** 内容指纹，与对应锚点的 digest 同源，用于内容寻址。 */
  digest: string;
  /** 归一化文本长度（不含内容本身，降低 IR 体积）。 */
  textLength: number;
}

/** 物理视图：给写入端看的「施工图纸」。 */
export interface PhysicalDocument {
  view: 'physical';
  /** 主文档部件名。 */
  mainPart: string;
  nodes: PhysicalNode[];
}

/** 对应表中的一条记录：一个语义节点 ↔ 一个或多个物理节点。 */
export interface SourceMapEntry {
  id: SemanticId;
  /** 语义节点种类，便于按类型检索。 */
  kind: string;
  /** 主物理指针（用于写入时的默认落点）。 */
  pointer: string;
  /** 全部物理指针（段内 run 被拆分时为 1:N）。 */
  pointers: string[];
  anchor: NodeAnchor;
}

/**
 * 对应哈希表 —— 双 IR 的「账本」。
 *
 * 三张索引各司其职，缺一不可：
 *   - bySemanticId : 正向查找（拿到语义 id，要改它 → 找到物理落点）；
 *   - byPointer    : 反向查找（拿到物理节点，要问它是什么 → 找到语义节点）；
 *   - byFingerprint: 内容寻址（拿一段文本，问「它对应哪些节点」）。
 *
 * 之所以用普通对象而不是 Map：Map 无法 JSON 序列化，会破坏「输出可序列化」这条
 * 硬性不变量，而对应表恰恰是必须跨越 RPC 边界的核心数据。
 */
export interface SourceMap {
  scheme: SourceMapScheme;
  entries: SourceMapEntry[];
  bySemanticId: Record<SemanticId, string[]>;
  byPointer: Record<string, SemanticId>;
  byFingerprint: Record<string, SemanticId[]>;
}

/** 双 IR 聚合体 —— `execute` 的核心产出。 */
export interface DocxDualIR {
  /** 视图数量，恒为 2（语义 + 物理）。 */
  viewCount: typeof DUAL_VIEW_COUNT;
  scheme: SourceMapScheme;
  semantic: SemanticDocument;
  physical: PhysicalDocument;
  sourceMap: SourceMap;
  /**
   * 表现层观察，按段落锚点索引；`parseFormatting` 关闭时整个字段缺席。
   *
   * 单独成节而不是塞进 `semantic.blocks`：语义视图回答「这一段说的是什么、
   * 它是谁」，本节回答「它长什么样」。需要按参照文档重建版面的消费方读这里，
   * 不读的消费方不为 run 级数据付费。`blockId` 是同一段的语义 id，可直接 join。
   */
  formatting?: {
    view: 'formatting';
    paragraphs: {
      blockId: string;
      pointer: string;
      alignment: string | null;
      /**
       * 有效首行缩进，**已解析样式链**：段落直接声明优先，否则沿 basedOn 找样式，
       * 再退到 docDefaults；谁都没声明就是 0（样式表读不到时整项缺席，因为那时
       * 「不知道」和「没有」不是一回事）。twips 与「百分之一字符」原样保留，由
       * 消费方按自己的字号口径换算。悬挂缩进不上报，它属于列表标记的排布。
       */
      indent?: { firstLineTwips?: number; firstLineChars?: number };
      runs: { text: string; bold?: boolean; size?: number; eastAsia?: string; color?: string }[];
    }[];
  };
}

/**
 * 本模块产出的 IR。
 *
 * 它【是】一份 office-core 的 `FormatIR`（因此 Profile 可以按通用方式消费），
 * 额外通过 `content` 字段承载双 IR。这样公共表面保持统一，
 * 而模块特有载荷不会污染 office-core 的定义。
 */
export interface DocxParseIR extends FormatIR {
  content: DocxDualIR;
}

/* -------------------------------------------------------------------------- */
/* 输入                                                                        */
/* -------------------------------------------------------------------------- */

/** 单次调用的可选项（在模块配置之上做细粒度覆盖）。 */
export interface DocxParseOptions {
  /** 调用方声称的扩展名（如 `docm`），用于类型不匹配检测。 */
  declaredExtension?: string;
  /** 调用方声称的媒体类型，用于类型不匹配检测。 */
  declaredMimeType?: string;
  /** 本次调用覆盖的能力开关。 */
  featureFlags?: Partial<FeatureFlags>;
  /** 本次调用覆盖的资源预算。 */
  limits?: Partial<LimitConfig>;
  /** 本次调用覆盖的超时时间。 */
  timeoutMs?: number;
  /** 为 true 时，`inspect` 顺带执行局部验证并附带报告。 */
  verifyAfterInspect?: boolean;
}

/** 三个接口共用的输入信封：可序列化、按请求隔离。 */
export interface DocxModuleInput {
  artifactRef: ArtifactRef;
  operation: DocxOperation;
  options?: DocxParseOptions;
  /** 请求 id，用于串联日志与遥测。 */
  requestId: string;
}

/** `inspect` 入参（可选携带策略以支持快速失败）。 */
export interface InspectInput extends DocxModuleInput {
  operation: 'inspect';
  policy?: SafetyPolicy;
}

/** `execute` 入参。 */
export interface ExecuteInput extends DocxModuleInput {
  operation: 'execute';
}

/** `verify` 入参；相较其他接口多一个安全策略。 */
export interface VerifyInput extends DocxModuleInput {
  operation: 'verify';
  policy: SafetyPolicy;
}

/* -------------------------------------------------------------------------- */
/* 输出                                                                        */
/* -------------------------------------------------------------------------- */

/** 单条验证检查项。 */
export interface VerificationCheck {
  /** 检查项 id（稳定，便于快照对比）。 */
  id: string;
  /** pass=通过；fail=不满足；skip=策略未声明该要求，无法判断。 */
  status: 'pass' | 'fail' | 'skip';
  /** 失败时对整体结果的影响级别。 */
  severity: 'error' | 'warn' | 'info';
  message: string;
  details?: JsonObject;
}

/** 验证结果统计。 */
export interface VerificationSummary {
  total: number;
  passed: number;
  failed: number;
  skipped: number;
}

/**
 * 本模块的局部验证报告。
 *
 * 它满足 office-test-kit 的形状契约，因此 Profile 可以直接把它喂给共享的报告链路。
 */
export interface VerificationReport extends VerificationReportLike {
  ok: boolean;
  partial: boolean;
  checks: VerificationCheck[];
  summary: VerificationSummary;
  /** 本次验证所用的策略 id。 */
  policyId?: string;
}

/** 单次调用的性能与规模摘要（不含任何文档内容）。 */
export interface TelemetrySummary {
  /** 解析脚本的协议版本，便于跨版本排查。 */
  parseVersion: number;
  /** 实际生效的 XML 后端（lxml 或 stdlib）。 */
  xmlBackend: string;
  /** 引擎耗时（毫秒）。 */
  engineMs: number;
  /** 端到端总耗时（毫秒）。 */
  totalMs: number;
  partCount: number;
  relationshipCount: number;
  blockCount: number;
  tableCount: number;
  annotationCount: number;
  sourceMapEntries: number;
}

/** `execute` 的结果载荷。 */
export interface ExecuteResult {
  ir: DocxParseIR;
  /** 补全后的 artifact 引用（回填媒体类型、大小与摘要）。 */
  artifact: ArtifactRef;
}

/** 所有接口共用的输出信封：结果 + 产物 + 告警 + 可选验证与遥测。 */
export interface DocxModuleOutput<TResult> {
  moduleId: DocxParseModuleId;
  requestId: string;
  operation: DocxOperation;
  result: TResult;
  /** 本模块产出的新 artifact 引用（当前为空：解析不产生新文件）。 */
  artifacts: ArtifactRef[];
  warnings: Warning[];
  verification?: VerificationReport;
  telemetry?: TelemetrySummary;
}

export type InspectOutput = DocxModuleOutput<FormatProfile>;
export type ExecuteOutput = DocxModuleOutput<ExecuteResult>;
export type VerifyOutput = DocxModuleOutput<VerificationReport>;

/* -------------------------------------------------------------------------- */
/* 遥测 / 日志端口                                                             */
/* -------------------------------------------------------------------------- */

/** 遥测级别。 */
export type TelemetryLevel = 'debug' | 'info' | 'warn' | 'error';

/** 单条遥测事件。所有字段都必须可 JSON 序列化。 */
export interface TelemetryEvent {
  /** 事件名，例如 `docx-parse.engine.parse`。 */
  name: string;
  level: TelemetryLevel;
  /** 事件发生的毫秒时间戳。 */
  timestampMs: number;
  /** 跨度耗时（仅结束事件有值）。 */
  durationMs?: number;
  /** 事件属性；禁止写入文档内容或绝对路径。 */
  attributes: Record<string, JsonValue>;
}

/** 遥测接收端。Profile 可注入自己的实现。 */
export interface TelemetrySink {
  record(event: TelemetryEvent): void;
}

/** 简单日志抽象（与遥测解耦，便于单测替换）。 */
export interface Logger {
  debug(message: string, attributes?: JsonObject): void;
  info(message: string, attributes?: JsonObject): void;
  warn(message: string, attributes?: JsonObject): void;
  error(message: string, attributes?: JsonObject): void;
}

/** 模块运行时上下文：配置 + 可选遥测/日志。 */
export interface ModuleContext {
  config: ModuleConfig;
  telemetry?: TelemetrySink;
  logger?: Logger;
}

/* -------------------------------------------------------------------------- */
/* 注册                                                                        */
/* -------------------------------------------------------------------------- */

export type DocxParseFn = (input: InspectInput) => Promise<InspectOutput>;
export type DocxExecuteFn = (input: ExecuteInput) => Promise<ExecuteOutput>;
export type DocxVerifyFn = (input: VerifyInput) => Promise<VerifyOutput>;

/** 三个接口的实现集合。 */
export interface DocxParseHandlers {
  inspect: DocxParseFn;
  execute: DocxExecuteFn;
  verify: DocxVerifyFn;
}

/** 依赖引用描述。 */
export interface ModuleDependencyRef {
  name: string;
  /** module=Profile 内模块；runtime=外部运行时（如 python 解释器）。 */
  kind: 'module' | 'runtime';
  optional?: boolean;
}

/** 与 `module.json` 一一对应的模块清单，供 Profile registry 消费。 */
export interface DocxParseModuleDefinition {
  id: DocxParseModuleId;
  version: string;
  /** 在 Profile 中的分组名。 */
  profileGroup: 'DOCX';
  summary: string;
  capabilities: readonly DocxParseCapability[];
  dependencies: readonly ModuleDependencyRef[];
  /** 描述 `ModuleConfig` 的 JSON-Schema 子集。 */
  configSchema: JsonObject;
}

/** 交付给 Profile 的模块实例。 */
export interface DocxParseModule {
  definition: DocxParseModuleDefinition;
  handlers: DocxParseHandlers;
  /** 释放模块资源（Python 桥接为无状态，故为空实现）。 */
  dispose(): Promise<void>;
}

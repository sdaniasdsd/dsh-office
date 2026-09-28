/**
 * docx-parse 公开契约 —— 【已冻结的表面】。
 *
 * 本文件导出的所有内容，都属于本模块在 Profile 中的「注册表面」。修改它们对下游
 * 消费者而言是破坏性变更，因此必须把它当作带版本号的接口来对待，而不是内部实现细节。
 *
 * 模块强制保证的三条不变量（由 `verifier.ts` 兜底检查）：
 *   1. 任何「输入」都可 JSON 序列化；
 *   2. 任何「输出」都可 JSON 序列化，且不包含任何底层引擎对象
 *      （不允许出现 Python 句柄、`zipfile`/`lxml` 实例、Buffer、类实例）；
 *   3. 错误分类、artifact 引用与验证报告由本模块自己拥有，
 *      底层库无法直接改写它们。
 *
 * 上游模块仅以 `import type` 方式消费，因此本模块与
 * office-core / office-safety / office-files 之间不存在运行时耦合。
 */
import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { VerificationReportLike } from 'office-test-kit';

import type { DocxParseIR } from './domain/docx-parse';

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

/**
 * 操作名与能力名同构。
 * 之所以单独起一个别名，是为了让「输入里携带的 operation 字段」在类型上自解释，
 * 而不是散落成裸字符串。
 */
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
 * 本模块只做结构解析，因此「没有标题」「样式缺失」等一律是告警而非错误——
 * 是否阻断由上层编排决定。
 *
 * 不变量：表里的每个码都必须有真实产出路径（引擎问题、mapper 直接产出，
 * 或 verifier/index 的判定）。保留「预留但无人发出」的码会让契约承诺大于
 * 实际能力，且不会有任何测试失败来提醒——因此宁可等到需要时再加一行。
 * 相关测试见 `tests/contract.test.ts` 的 warning-code 可见性用例。
 */
export const DOCX_PARSE_WARNING_CODES = {
  /** 结构合法但没有任何可见内容。 */
  EMPTY_DOCUMENT: 'EMPTY_DOCUMENT',
  /** 文档没有任何标题。 */
  NO_HEADINGS: 'NO_HEADINGS',
  /** 缺少 `word/styles.xml`。 */
  STYLES_MISSING: 'STYLES_MISSING',
  /** 段落引用了 styles.xml 中不存在的样式 id。 */
  BROKEN_STYLE_REFERENCE: 'BROKEN_STYLE_REFERENCE',
  /** 内部关系的目标部件在包内不存在。 */
  DANGLING_RELATIONSHIP: 'DANGLING_RELATIONSHIP',
  /** 文档包含批注。 */
  COMMENTS_PRESENT: 'COMMENTS_PRESENT',
  /** 文档包含脚注/尾注。 */
  FOOTNOTES_PRESENT: 'FOOTNOTES_PRESENT',
  /** 正文里存在本模块不解析的元素，其内容已被跳过。 */
  UNKNOWN_BLOCK_SKIPPED: 'UNKNOWN_BLOCK_SKIPPED',
  DECLARED_HINT_MISMATCH: 'DECLARED_HINT_MISMATCH',
  LIMIT_APPLIED: 'LIMIT_APPLIED',
  /** 仅完成部分解析（超限、XML 损坏或能力开关关闭）。 */
  PARTIALLY_PARSED: 'PARTIALLY_PARSED',
  XML_BACKEND_FALLBACK: 'XML_BACKEND_FALLBACK',
} as const;

/** 告警码联合类型。 */
export type DocxParseWarningCode = keyof typeof DOCX_PARSE_WARNING_CODES;

/* -------------------------------------------------------------------------- */
/* 配置（由 Profile 注入；模块从不读取全局配置）                                */
/* -------------------------------------------------------------------------- */

/**
 * 深度引擎配置。
 *
 * 用途：为 `execute` / `verify` 换用解析质量更高的后端（当前为 Docling）。
 * 之所以与轻量引擎分开配置，是因为二者在成本上差一个量级：
 * 深度引擎要加载模型，启动即数秒、占用数百 MB，因此【绝不】允许进入 `inspect`——
 * 「先看一眼文档长什么样」必须保持秒级响应。
 */
export interface DeepEngineConfig {
  /** 深度引擎族。 */
  driver: 'docling';
  /** 运行深度引擎所用的解释器；省略时复用 `EngineConfig.pythonPath`。 */
  pythonPath?: string;
  /** 桥接脚本绝对路径；省略时取 adapter.ts 同目录下的 docx_parse_docling.py。 */
  scriptPath?: string;
  /** 传给子进程的额外环境变量。 */
  env?: Record<string, string>;
  /**
   * 深度引擎独立的超时（毫秒）。
   *
   * 必要性：模型加载本身就可能超过轻量引擎的 `timeoutMs`，
   * 用同一个预算会把「正常但慢」误判成 `ENGINE_TIMEOUT`。
   */
  timeoutMs?: number;
}

/** 引擎相关配置。 */
export interface EngineConfig {
  /** 引擎族：默认的轻量解析器。`inspect` 恒使用它。 */
  driver: 'python';
  /** 启动解析脚本所用的解释器（可为命令名或绝对路径）。 */
  pythonPath: string;
  /** 解析脚本的绝对路径；省略时取 adapter.ts 同目录下的 docx_parse.py。 */
  scriptPath?: string;
  /** 传给子进程的额外环境变量。 */
  env?: Record<string, string>;
  /**
   * 可选的深度引擎。
   *
   * 省略时三条路径共用 `driver`（即当前的纯标准库行为，保持不变）。
   * 一旦配置，`execute` 与 `verify` 改走深度引擎，`inspect` 依旧不变。
   * 无论走哪个引擎，公开契约、错误分类、artifact 引用与验证报告都不受影响。
   */
  deep?: DeepEngineConfig;
}

/** 资源预算，用于抵御 zip 炸弹与超大 OOXML 包。 */
export interface LimitConfig {
  /** 检查的归档条目数上限，超出即中止。 */
  maxArchiveEntries: number;
  /** 单个条目解压后的大小上限。 */
  maxEntryUncompressedBytes: number;
  /** 所有条目解压后的总大小上限。 */
  maxTotalUncompressedBytes: number;
  /** 解析的关系条数上限。 */
  maxRelationships: number;
  /** 正文块（段落 + 表格）数量上限。 */
  maxBlocks: number;
  /** 表格单元格总数上限。 */
  maxTableCells: number;
  /** 样式条数上限。 */
  maxStyles: number;
  /** 批注条数上限。 */
  maxComments: number;
  /** 脚注/尾注条数上限。 */
  maxFootnotes: number;
}

/** 能力开关。关闭某项后引擎不再解析对应结构。 */
export interface FeatureFlags {
  parseParagraphs: boolean;
  parseHeadings: boolean;
  parseTables: boolean;
  parseStyles: boolean;
  parseRelationships: boolean;
  parseComments: boolean;
  parseFootnotes: boolean;
  /** 是否解析 `docProps/core.xml` 元数据。 */
  extractMetadata: boolean;
  /** 为 false 时，引擎只上报超限但不中止（用于「尽力而为」的宽限场景）。 */
  enforceLimits: boolean;
}

/** 模块完整配置：由 Profile 注入，模块不读全局配置。 */
export interface ModuleConfig {
  engine: EngineConfig;
  runtimeRoot?: string;
  runtimePackageNames?: readonly string[];
  limits: LimitConfig;
  timeoutMs: number;
  featureFlags: FeatureFlags;
}

/* -------------------------------------------------------------------------- */
/* 策略（verify 用；复用 office-safety 的策略标识与加密处置）                    */
/* -------------------------------------------------------------------------- */

/**
 * docx-parse 的校验策略。
 *
 * 与 docx-inspect 的安全策略不同：本模块校验的是「结构质量」而非「安全指标」。
 * 因此这里只复用 office-safety 中与本模块相关的两个字段（id / allowEncrypted），
 * 其余是解析质量要求。所有开关都是可选的：缺省表示「本策略未声明该要求」，
 * 对应检查会被标记为 `skip` 而不是 `pass`。
 */
export interface ParsePolicy extends Pick<SafetyPolicy, 'id' | 'allowEncrypted'> {
  /** 是否要求文档包含至少一个标题。 */
  requireHeadings?: boolean;
  /** 是否要求文档至少包含 N 个非空段落。 */
  minParagraphs?: number;
  /** 是否要求存在样式表。 */
  requireStyles?: boolean;
  /** 是否要求所有段落引用的样式都真实存在。 */
  requireResolvedStyles?: boolean;
  /** 是否要求所有内部关系都有对应部件。 */
  requireNoDanglingRelationships?: boolean;
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

/**
 * `inspect` 入参。
 *
 * 携带可选 `policy`：spec 规定 `inspect` 的失败语义包含「安全策略拒绝」，
 * 因此当调用方提供了策略时，`inspect` 会在解析后立即评估它，
 * 命中 error 级检查即快速失败（返回 SAFETY_POLICY_DENIED）。
 */
export interface InspectInput extends DocxModuleInput {
  operation: 'inspect';
  /** 可选：提供后 `inspect` 会顺带做策略检查并可能快速失败。 */
  policy?: ParsePolicy;
}

/** `execute` 入参。 */
export interface ExecuteInput extends DocxModuleInput {
  operation: 'execute';
}

/** `verify` 入参；相较其他接口多一个策略。 */
export interface VerifyInput extends DocxModuleInput {
  operation: 'verify';
  policy: ParsePolicy;
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
 * 它满足 office-test-kit 的形状契约，因此 Profile 可以直接把它喂给共享的报告链路，
 * 无需再做一次格式转换。
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
  /**
   * 本次实际执行的引擎名。
   *
   * 必要性：`execute` / `verify` 可能走深度引擎而 `inspect` 恒走轻量引擎，
   * 不记录引擎名就无法解释同一文档在不同接口上的耗时差异。
   */
  engine: string;
  /** 解析脚本的协议版本，便于跨版本排查。 */
  parseVersion: number;
  /** 实际生效的 XML 后端（lxml 或 stdlib）。 */
  xmlBackend: string;
  /** 引擎耗时（毫秒）。 */
  engineMs: number;
  /** 端到端总耗时（毫秒）。 */
  totalMs: number;
  blockCount: number;
  tableCount: number;
  styleCount: number;
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

/**
 * 最小化 registry 端口。
 *
 * 真实实现由 Profile 提供；本模块只需要「把我注册进去」这一件事，
 * 从而保持自身不含任何编排逻辑。
 */
export interface ProfileRegistry {
  registerModule(module: DocxParseModule): void | Promise<void>;
}

/* -------------------------------------------------------------------------- */
/* 领域类型再导出（单一数据源，避免「双 IR」）                                    */
/* -------------------------------------------------------------------------- */

export type {
  Alignment,
  Block,
  CommentRecord,
  ContainerKind,
  DocumentCounts,
  DocumentKind,
  DocumentMetadata,
  DocumentVariant,
  DocxParseIR,
  FootnoteRecord,
  OutlineEntry,
  ParagraphRecord,
  ParseIssue,
  ParseResult,
  RelationshipRecord,
  RunRecord,
  StyleRecord,
  StyleType,
  TableCell,
  TableRecord,
} from './domain/docx-parse';

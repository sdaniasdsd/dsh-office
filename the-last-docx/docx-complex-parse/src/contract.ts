/**
 * docx-complex-parse 公开契约 —— 【已冻结的表面】。
 *
 * 本文件导出的所有内容，都属于本模块在 Profile 中的「注册表面」。修改它们对下游
 * 消费者而言是破坏性变更，因此必须把它当作带版本号的接口来对待。
 *
 * 本模块负责【复杂 DOCX 解析】，核心处理范围五项：
 *
 *   1. 复杂表格 —— 合并单元格还原为逻辑网格、跨页表头、行列跨度；
 *   2. 跨页关系 —— 显式分页符、渲染分页标记、分节与页面尺寸变化；
 *   3. 浮动对象 —— 图片、图表、图形、文本框、OLE 及其锚定与环绕方式；
 *   4. 来源坐标 —— 每个结构既报 XML 结构位置，也报页内几何位置（拿得到时）；
 *   5. 置信度   —— 每项结论都标注它有多可信、依据是什么。
 *
 * 三条必须守住的纪律：
 *   1. 任何输入输出都可 JSON 序列化，且不包含任何底层引擎对象
 *      （Docling / Mammoth / MarkItDown 的返回对象一律不得逃出 mapper）；
 *   2. 错误分类、artifact 引用、验证报告由本模块自己拥有，底层库无权改写；
 *   3. 引擎可替换而不改动本文件——上层只看见这份契约。
 *
 * 职责边界：本模块【只读】。把改动写回 OOXML 属于另一个模块的范围。
 *
 * 本文件刻意自包含：除 office-core / office-safety / office-test-kit 三个公共契约
 * 外，不引用任何兄弟模块的内部类型。
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
export const DOCX_COMPLEX_PARSE_MODULE_ID = 'docx-complex-parse' as const;
export type DocxComplexParseModuleId = typeof DOCX_COMPLEX_PARSE_MODULE_ID;

/** 本模块对外发布的三个接口。 */
export const DOCX_COMPLEX_PARSE_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export type DocxComplexParseCapability = (typeof DOCX_COMPLEX_PARSE_CAPABILITIES)[number];

/** 操作名与能力名同构（让输入里的 operation 字段在类型上自解释）。 */
export type DocxComplexOperation = DocxComplexParseCapability;

/** 本模块产出的复杂结构方案名。 */
export const COMPLEX_LAYOUT_SCHEME = 'complex-layout-v1' as const;
export type ComplexLayoutScheme = typeof COMPLEX_LAYOUT_SCHEME;

/** 复杂结构的稳定标识。 */
export type ComplexNodeId = string;

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
export const DOCX_COMPLEX_PARSE_ERROR_CODES = {
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
  /** 无法启动引擎（Docling / Mammoth / MarkItDown 不可用）。 */
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
   * 【本模块独有】版面信息不可用：需要页面几何或分页关系，但当前引擎无法产出。
   *
   * 与 `PARSE_FAILED` 的区别很重要——文档没坏，只是「页」这个概念需要排版模型，
   * 而所选引擎给不出。调用方可以选择降级消费（structure 坐标仍然有效），
   * 也可以视其为硬失败。由 `featureFlags.requirePageGeometry` 决定走哪条路。
   */
  LAYOUT_UNAVAILABLE: 'LAYOUT_UNAVAILABLE',
  /** `verify` 无法产出报告，或某项强制检查硬失败。 */
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;

/** 错误码联合类型。 */
export type DocxComplexParseErrorCode = keyof typeof DOCX_COMPLEX_PARSE_ERROR_CODES;

/* -------------------------------------------------------------------------- */
/* 告警码                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 告警码表。
 *
 * 注意区分：错误码表示「调用失败」，告警码表示「调用成功但结果需要被注意」。
 * 复杂结构解析天然是「尽力而为」的——遇到读不懂或读不全的结构时降级并告警，
 * 而不是让整次调用失败。
 */
export const DOCX_COMPLEX_PARSE_WARNING_CODES = {
  /** 文档结构合法，但正文中没有任何可见块。 */
  EMPTY_DOCUMENT: 'EMPTY_DOCUMENT',
  /** 只解析了部分内容（存在跳过或未建模的元素）。 */
  PARTIALLY_PARSED: 'PARTIALLY_PARSED',
  /** 遇到本模块未建模的 OOXML 元素。 */
  UNSUPPORTED_CONTENT: 'UNSUPPORTED_CONTENT',
  /** 命中资源预算但被宽免（`enforceLimits` 关闭）。 */
  LIMIT_APPLIED: 'LIMIT_APPLIED',
  /** 嵌套表格被压平（内层结构未展开）。 */
  NESTED_TABLE_FLATTENED: 'NESTED_TABLE_FLATTENED',
  /** 合并区域自相矛盾（重叠或越界），已按最保守方式解释。 */
  MERGE_INCONSISTENT: 'MERGE_INCONSISTENT',
  /** 页面几何拿不到，坐标降级为结构位置。 */
  PAGE_GEOMETRY_UNAVAILABLE: 'PAGE_GEOMETRY_UNAVAILABLE',
  /** 分页位置是估算的（由启发式或渲染标记推导，非权威）。 */
  PAGE_BREAK_ESTIMATED: 'PAGE_BREAK_ESTIMATED',
  /** 浮动对象的锚定关系只能近似（锚点信息不完整）。 */
  FLOAT_ANCHOR_APPROXIMATED: 'FLOAT_ANCHOR_APPROXIMATED',
  /** 某项结论的置信度低于 `config.confidenceFloor`。 */
  LOW_CONFIDENCE: 'LOW_CONFIDENCE',
  /** 产物已加密，无法解析内容。 */
  ENCRYPTED_ARTIFACT: 'ENCRYPTED_ARTIFACT',
  /** 验证未通过（作为告警广播，便于只消费 warnings 的上层感知）。 */
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;

/** 告警码联合类型。 */
export type DocxComplexParseWarningCode = keyof typeof DOCX_COMPLEX_PARSE_WARNING_CODES;

/* -------------------------------------------------------------------------- */
/* 配置（由 Profile 注入；模块从不读取全局配置）                                */
/* -------------------------------------------------------------------------- */

/**
 * 本模块支持的引擎族。
 *
 * `rdocx` 是本模块的内置引擎：它用自带的 OOXML 读取器出结构、用 rdocx 出页几何，
 * 不依赖外部转换工具。其余三个是候选的外部转换引擎，尚未接入——若宿主显式选了
 * 它们，`execute` 会以 `ENGINE_UNAVAILABLE` 快速失败，而不是静默换用别的引擎。
 */
export const COMPLEX_ENGINE_DRIVERS = ['rdocx', 'docling', 'mammoth', 'markitdown'] as const;
export type ComplexEngineDriver = (typeof COMPLEX_ENGINE_DRIVERS)[number];

/** 本模块实际内置了实现、可直接使用的引擎。 */
export const IMPLEMENTED_ENGINE_DRIVERS = ['rdocx'] as const satisfies readonly ComplexEngineDriver[];

/** 引擎相关配置。 */
export interface EngineConfig {
  /** 引擎族。三者能力侧重不同，见 README 的对照。 */
  driver: ComplexEngineDriver;
  /** 启动引擎所用的解释器/命令（可为命令名或绝对路径）。 */
  pythonPath: string;
  /** 引擎脚本或模块的路径；省略时取 adapter 的默认实现。 */
  scriptPath?: string;
  /** 传给子进程的额外环境变量。 */
  env?: Record<string, string>;
}

/** 资源预算，用于抵御 zip 炸弹与超大 OOXML 包。 */
export interface LimitConfig {
  /** 检查的归档条目数上限，超出即中止。 */
  maxArchiveEntries: number;
  /** 单个条目解压后字节数上限。 */
  maxEntryUncompressedBytes: number;
  /** 全部条目解压后总字节数上限。 */
  maxTotalUncompressedBytes: number;
  /** 关系条目数上限。 */
  maxRelationships: number;
  /** 正文块数量上限。 */
  maxBlocks: number;
  /** 单表单元格数量上限。 */
  maxTableCells: number;
  /** 浮动对象数量上限。 */
  maxFloatingObjects: number;
  /** 页数上限（仅在有版面信息时生效）。 */
  maxPages: number;
}

/** 能力开关。每一项都对应本模块五项处理范围中的一部分。 */
export interface FeatureFlags {
  /** 解析表格并还原逻辑网格。 */
  parseTables: boolean;
  /** 解析合并区域（gridSpan / vMerge）。 */
  parseMerges: boolean;
  /** 解析显式分页符与渲染分页标记。 */
  parsePageBreaks: boolean;
  /** 解析分节（页面尺寸、方向、分栏变化）。 */
  parseSections: boolean;
  /** 解析浮动对象。 */
  parseFloatingObjects: boolean;
  /** 采集页面几何坐标（需要版面分析引擎）。 */
  parsePageGeometry: boolean;
  /** 为每项结论计算置信度。 */
  computeConfidence: boolean;
  /** 页面几何不可用时是否硬失败（false 则降级为结构坐标并告警）。 */
  requirePageGeometry: boolean;
  /** 是否执行资源预算检查。 */
  enforceLimits: boolean;
}

/** 模块配置：由 Profile 注入，模块不读取全局配置。 */
export interface ModuleConfig {
  engine: EngineConfig;
  limits: LimitConfig;
  timeoutMs: number;
  featureFlags: FeatureFlags;
  /** 置信度下限：低于此值的结论会附带 LOW_CONFIDENCE 告警。 */
  confidenceFloor: number;
}

/* -------------------------------------------------------------------------- */
/* 处理范围一：来源坐标                                                         */
/* -------------------------------------------------------------------------- */

/** 坐标单位。EMU 是 OOXML 原生单位（914400 EMU = 1 英寸）。 */
export type CoordinateUnit = 'emu' | 'pt' | 'px' | 'twip';

/** 页内矩形。原点在页面左上角，y 轴向下。 */
export interface PageBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * 来源坐标 —— 一个结构「在哪儿」的完整回答。
 *
 * 刻意把两种坐标分开记录，因为它们来自不同的证据源、可靠性也不同：
 *   - `structuralPath`：XML 结构位置，纯解析即可得到，永远存在；
 *   - `pageIndex` / `box`：页内几何位置，需要版面模型，可能为 null。
 *
 * 把 `pageIndex` 设为 `number | null` 而不是 `-1` 之类的哨兵值：读的人必须
 * 显式处理「没有页面信息」这一情形，而不是不知不觉地把它当成第 -1 页。
 */
export interface SourceCoordinate {
  /** 承载该结构的部件名，例如 `word/document.xml`。 */
  part: string;
  /** 部件内的 XPath 风格结构路径，例如 `/w:document/w:body/w:tbl[1]`。 */
  structuralPath: string;
  /** 页序号（0 基）；无版面信息时为 null。 */
  pageIndex: number | null;
  /** 页内矩形；无版面信息时为 null。 */
  box: PageBox | null;
  /** 页内坐标的单位；`box` 为 null 时同样为 null。 */
  unit: CoordinateUnit | null;
}

/* -------------------------------------------------------------------------- */
/* 处理范围二：置信度                                                           */
/* -------------------------------------------------------------------------- */

/**
 * 置信度的依据来源，按可信度从高到低排列。
 *
 * 这个字段比 `score` 本身更重要：一个 0.6 分的结论，如果依据是 `native`
 * （OOXML 里明确写着），和依据是 `heuristic`（我们猜的），补救方式完全不同。
 */
export type ConfidenceSource =
  /** 文档里明确声明的事实（如 w:gridSpan="2"）。 */
  | 'native'
  /** 由结构关系推导（如由 tblGrid 列数对齐）。 */
  | 'structural'
  /** 由启发式规则推断（如按缩进猜标题层级）。 */
  | 'heuristic'
  /** 由版面模型/ML 给出（如 Docling 的版面识别）。 */
  | 'model';

/** 一项结论的置信度。 */
export interface Confidence {
  /** 0..1；1 表示文档中明确声明，无需推断。 */
  score: number;
  /** 依据来源。 */
  source: ConfidenceSource;
  /** 判定理由（人类可读，用于解释「为什么是这个分」）。 */
  notes: string[];
}

/* -------------------------------------------------------------------------- */
/* 处理范围三：复杂表格                                                         */
/* -------------------------------------------------------------------------- */

/** 逻辑网格中的一个格子。 */
export interface GridCell {
  /** 行序号（0 基，逻辑网格坐标）。 */
  row: number;
  /** 列序号（0 基，逻辑网格坐标）。 */
  column: number;
  /** 单元格文本。被合并覆盖的占位格为空字符串。 */
  text: string;
  /**
   * 该格由哪个真实单元格承载。
   * `null` 表示这一格被左侧/上方的合并区域覆盖，本身不是独立单元格。
   */
  origin: { row: number; column: number } | null;
  /** 真实单元格自身的坐标；占位格为 null。 */
  coordinate: SourceCoordinate | null;
  /** 文本与归属判定的置信度。 */
  confidence: Confidence;
}

/** 合并区域的种类。 */
export type MergeKind = 'gridSpan' | 'vMerge' | 'both';

/** 一处合并区域（已还原为逻辑网格坐标）。 */
export interface MergeRegion {
  /** 起始行（0 基）。 */
  startRow: number;
  /** 起始列（0 基）。 */
  startColumn: number;
  /** 纵向跨度，至少 1。 */
  rowSpan: number;
  /** 横向跨度，至少 1。 */
  columnSpan: number;
  kind: MergeKind;
  /** 被合并格的原生文本（可能为空）。 */
  text: string;
}

/**
 * 复杂表格 —— 合并已还原为逻辑网格。
 *
 * `grid` 是展开后的矩形矩阵（行数 × 列数），因此可以直接按 `grid[row][column]`
 * 取值，无需调用方自己处理合并。`merges` 保留合并区域的原始信息，
 * 供需要还原成 OOXML 形态的下游使用。
 */
export interface ComplexTable {
  id: ComplexNodeId;
  /** 表格自身的坐标。 */
  coordinate: SourceCoordinate;
  confidence: Confidence;
  /** 逻辑列数（以 tblGrid 为准，缺失时按最大行宽推导）。 */
  columnCount: number;
  /** 逻辑行数。 */
  rowCount: number;
  /** 展开后的逻辑网格，维度恒为 rowCount × columnCount。 */
  grid: GridCell[][];
  /** 全部合并区域。 */
  merges: MergeRegion[];
  /** 重复表头行数（w:tblHeader）；无表头为 0。 */
  headerRowCount: number;
  /** 表格是否跨页（依据分页信息推导，无版面信息时为 null）。 */
  spansPages: boolean | null;
  /** 表格占据的页序号区间 `[起, 止]`（含端点）；无版面信息时为 null。 */
  pageRange: [number, number] | null;
}

/* -------------------------------------------------------------------------- */
/* 处理范围四：跨页关系                                                         */
/* -------------------------------------------------------------------------- */

/**
 * 分页类型 —— 说明「这个分页是怎么来的」，可信度依次递减。
 */
export type PageBreakKind =
  /** 文档中显式的分页符（w:br w:type="page"）。 */
  | 'explicit'
  /** Word 渲染后写回的分页标记（w:lastRenderedPageBreak）。 */
  | 'rendered'
  /** 分节符导致的分页（sectPr + type）。 */
  | 'section'
  /** 由版面模型推断。 */
  | 'model';

/** 一处分页。 */
export interface PageBreak {
  kind: PageBreakKind;
  /** 分页发生在哪个结构之前（该结构是新一页的第一个内容）。 */
  beforeNodeId: ComplexNodeId | null;
  /** 分页前的页序号（0 基）；无版面信息时为 null。 */
  pageIndexBefore: number | null;
  /** 分页后的页序号（0 基）；无版面信息时为 null。 */
  pageIndexAfter: number | null;
  coordinate: SourceCoordinate;
  confidence: Confidence;
}

/** 页尺寸与版式。 */
export interface PageGeometry {
  /** 页宽。 */
  width: number;
  /** 页高。 */
  height: number;
  unit: CoordinateUnit;
  /** 方向。 */
  orientation: 'portrait' | 'landscape';
  /** 页边距（上右下左）。 */
  margins: { top: number; right: number; bottom: number; left: number };
}

/** 一节（w:sectPr 对应的一段排版区间）。 */
export interface SectionInfo {
  /** 节序号（0 基）。 */
  index: number;
  /** 本节起始于哪个结构之前。 */
  startBeforeNodeId: ComplexNodeId | null;
  /** 本节页尺寸；无法判定时为 null。 */
  geometry: PageGeometry | null;
  /** 分栏数。 */
  columnCount: number;
  /** 节的开始方式。 */
  startType: 'continuous' | 'nextPage' | 'evenPage' | 'oddPage' | 'unknown';
}

/* -------------------------------------------------------------------------- */
/* 处理范围五：浮动对象                                                         */
/* -------------------------------------------------------------------------- */

/** 浮动对象的种类。 */
export type FloatingKind =
  | 'image'
  | 'chart'
  | 'shape'
  | 'textbox'
  | 'equation'
  | 'ole'
  | 'unknown';

/** 文字环绕方式。 */
export type WrapMode =
  | 'inline'
  | 'square'
  | 'tight'
  | 'through'
  | 'topAndBottom'
  | 'behindText'
  | 'inFrontOfText'
  | 'none';

/**
 * 浮动对象的锚定关系。
 *
 * 这是「浮动对象」区别于「内联对象」的关键：它不占正文流位置，
 * 而是相对某个基准（页、页边距、段落、字符）偏移定位。
 */
export interface FloatAnchor {
  /** 水平基准，例如 `page` / `margin` / `column` / `character`。 */
  relativeFromHorizontal: string;
  /** 垂直基准，例如 `page` / `margin` / `paragraph` / `line`。 */
  relativeFromVertical: string;
  /** 水平对齐方式（与 offsetX 二选一）。 */
  alignHorizontal: 'left' | 'center' | 'right' | null;
  /** 垂直对齐方式（与 offsetY 二选一）。 */
  alignVertical: 'top' | 'center' | 'bottom' | null;
  /** 相对基准的水平偏移；未声明时为 null。 */
  offsetX: number | null;
  /** 相对基准的垂直偏移；未声明时为 null。 */
  offsetY: number | null;
  /** 偏移量单位。 */
  unit: CoordinateUnit | null;
  /** 环绕方式。 */
  wrap: WrapMode;
  /** 是否位于文字下方。 */
  behindText: boolean;
}

/**
 * 浮动对象 —— 不占正文流位置的图形/媒体。
 *
 * 注意 `anchor.offsetX/offsetY` 是【相对基准的偏移】，不是绝对页面坐标。
 * 要得到绝对页面坐标，需要版面模型把基准位置解出来，结果落在 `coordinate.box`。
 * 二者都保留：前者是文档里写的事实，后者是推导结果。
 */
export interface FloatingObject {
  id: ComplexNodeId;
  kind: FloatingKind;
  /** 对象自身的坐标。 */
  coordinate: SourceCoordinate;
  confidence: Confidence;
  anchor: FloatAnchor;
  /** 对象尺寸；未声明时为 null。 */
  width: number | null;
  height: number | null;
  /** 尺寸单位。 */
  unit: CoordinateUnit | null;
  /** 承载该对象的部件（通常为 `word/document.xml`）。 */
  part: string;
  /** 目标媒体的关系 id（图片/图表/OLE 指向的关系）。 */
  relationshipId: string | null;
  /** 对象名称（`wp:docPr` 的 name）；未声明时为 null。 */
  name: string | null;
  /** 替代文本（`wp:docPr` 的 descr）；未声明时为 null。 */
  altText: string | null;
}

/* -------------------------------------------------------------------------- */
/* 复杂结构载荷                                                                 */
/* -------------------------------------------------------------------------- */

/** 正文块的种类。 */
export type ComplexBlockKind = 'table' | 'paragraph' | 'unknown';

/** 正文块 —— 本模块只对表格做深度建模，段落保持轻量。 */
export interface ComplexBlock {
  id: ComplexNodeId;
  kind: ComplexBlockKind;
  /** 纯文本（段落为正文；表格为按行拼接的摘要）。 */
  text: string;
  coordinate: SourceCoordinate;
  confidence: Confidence;
  /** 表格块在此指向 ComplexTable；非表格为 null。 */
  table: ComplexTable | null;
  /** All model fragments, including continuation pages. Absent in older producers. */
  fragments?: SourceCoordinate[];
}

/** 一页的汇总。 */
export interface PageInfo {
  /** 页序号（0 基）。 */
  index: number;
  /** 页尺寸；无法判定时为 null。 */
  geometry: PageGeometry | null;
  /** 本页包含的结构 id，按文档顺序。 */
  nodeIds: ComplexNodeId[];
  /** 本页是否由版面模型推断（而非文档声明）。 */
  inferred: boolean;
}

/**
 * 复杂结构 —— `execute` 的核心载荷。
 *
 * 五个数组/集合与五项处理范围一一对应，方便下游按需取用：
 *   blocks ↔ 复杂表格（经 block.table 访问）
 *   breaks / sections ↔ 跨页关系
 *   floats  ↔ 浮动对象
 *   coordinate ↔ 来源坐标（分布在每个结构上）
 *   confidence ↔ 置信度（分布在每个结构上）
 */
export interface ComplexContent {
  scheme: ComplexLayoutScheme;
  /** 包内主文档部件名。 */
  mainPart: string;
  /** 页汇总；无版面信息时为空数组。 */
  pages: PageInfo[];
  /** 正文块（按文档顺序，与物理顺序一致）。 */
  blocks: ComplexBlock[];
  /** 全部分页点（按文档顺序）。 */
  breaks: PageBreak[];
  /** 全部分节（按文档顺序）。 */
  sections: SectionInfo[];
  /** 全部浮动对象（按文档顺序）。 */
  floats: FloatingObject[];
  /** Main-body observation coverage; absent in older producers. Empty arrays alone do not prove absence. */
  coverage?: Record<'breaks' | 'sections' | 'floats', 'observed' | 'partial' | 'disabled' | 'unavailable'>;
}

/**
 * 本模块产出的 IR。
 *
 * 它【是】一份 office-core 的 `FormatIR`（因此 Profile 可以按通用方式消费），
 * 额外通过 `content` 字段承载复杂结构。这样公共表面保持统一，
 * 而模块特有载荷不会污染 office-core 的定义。
 */
export interface DocxComplexParseIR extends FormatIR {
  content: ComplexContent;
}

/* -------------------------------------------------------------------------- */
/* 输入                                                                        */
/* -------------------------------------------------------------------------- */

/** 单次调用的可选项（在模块配置之上做细粒度覆盖）。 */
export interface DocxComplexParseOptions {
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
  /** 本次调用覆盖的置信度下限。 */
  confidenceFloor?: number;
  /** 为 true 时，`inspect` 顺带执行局部验证并附带报告。 */
  verifyAfterInspect?: boolean;
}

/** 三个接口共用的输入信封：可序列化、按请求隔离。 */
export interface DocxComplexModuleInput {
  artifactRef: ArtifactRef;
  operation: DocxComplexOperation;
  options?: DocxComplexParseOptions;
  /** 请求 id，用于串联日志与遥测。 */
  requestId: string;
}

/** `inspect` 入参（可选携带策略以支持快速失败）。 */
export interface InspectInput extends DocxComplexModuleInput {
  operation: 'inspect';
  policy?: SafetyPolicy;
}

/** `execute` 入参。 */
export interface ExecuteInput extends DocxComplexModuleInput {
  operation: 'execute';
}

/** `verify` 入参；相较其他接口多一个安全策略。 */
export interface VerifyInput extends DocxComplexModuleInput {
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
  /** 解析协议的版本，便于跨版本排查。 */
  parseVersion: number;
  /** 实际生效的引擎族。 */
  engineDriver: ComplexEngineDriver;
  /** 引擎耗时（毫秒）。 */
  engineMs: number;
  /** 端到端总耗时（毫秒）。 */
  totalMs: number;
  partCount: number;
  relationshipCount: number;
  blockCount: number;
  tableCount: number;
  floatingObjectCount: number;
  pageCount: number;
  /** 页面几何是否可用（决定坐标精度）。 */
  pageGeometryAvailable: boolean;
}

/** `execute` 的结果载荷。 */
export interface ExecuteResult {
  ir: DocxComplexParseIR;
  /** 补全后的 artifact 引用（回填媒体类型、大小与摘要）。 */
  artifact: ArtifactRef;
}

/** 所有接口共用的输出信封：结果 + 产物 + 告警 + 可选验证与遥测。 */
export interface DocxComplexModuleOutput<TResult> {
  moduleId: DocxComplexParseModuleId;
  requestId: string;
  operation: DocxComplexOperation;
  result: TResult;
  /** 本模块产出的新 artifact 引用（当前为空：解析不产生新文件）。 */
  artifacts: ArtifactRef[];
  warnings: Warning[];
  verification?: VerificationReport;
  telemetry?: TelemetrySummary;
}

/** 三个接口各自的输出类型。 */
export type InspectOutput = DocxComplexModuleOutput<FormatProfile>;
export type ExecuteOutput = DocxComplexModuleOutput<ExecuteResult>;
export type VerifyOutput = DocxComplexModuleOutput<VerificationReport>;

/* -------------------------------------------------------------------------- */
/* 遥测                                                                        */
/* -------------------------------------------------------------------------- */

/** 遥测级别。 */
export type TelemetryLevel = 'debug' | 'info' | 'warn' | 'error';

/** 一条遥测事件。 */
export interface TelemetryEvent {
  level: TelemetryLevel;
  /** 事件名，例如 `docx-complex-parse.execute.failed`。 */
  name: string;
  /** 结构化上下文（不得包含文档正文或绝对路径）。 */
  attributes: Record<string, JsonValue>;
}

/** 遥测接收端；未注入时模块静默丢弃。 */
export interface TelemetrySink {
  emit(event: TelemetryEvent): void;
}

/** 模块内部使用的日志器。 */
export interface Logger {
  debug(name: string, attributes?: Record<string, JsonValue>): void;
  info(name: string, attributes?: Record<string, JsonValue>): void;
  warn(name: string, attributes?: Record<string, JsonValue>): void;
  error(name: string, attributes?: Record<string, JsonValue>): void;
}

/** 调用上下文：配置 + 可选遥测/日志。 */
export interface ModuleContext {
  config: ModuleConfig;
  telemetry?: TelemetrySink;
  logger?: Logger;
}

/* -------------------------------------------------------------------------- */
/* 模块定义                                                                    */
/* -------------------------------------------------------------------------- */

/** 三个 handler 的函数签名。 */
export type DocxComplexInspectFn = (input: InspectInput) => Promise<InspectOutput>;
export type DocxComplexExecuteFn = (input: ExecuteInput) => Promise<ExecuteOutput>;
export type DocxComplexVerifyFn = (input: VerifyInput) => Promise<VerifyOutput>;

/** 模块暴露给 Profile 的 handler 集合。 */
export interface DocxComplexParseHandlers {
  inspect: DocxComplexInspectFn;
  execute: DocxComplexExecuteFn;
  verify: DocxComplexVerifyFn;
}

/** 依赖声明项。 */
export interface ModuleDependencyRef {
  name: string;
  kind: 'module' | 'runtime';
  optional?: boolean;
}

/** Profile 注册时读取的模块清单。 */
export interface DocxComplexParseModuleDefinition {
  id: DocxComplexParseModuleId;
  version: string;
  profileGroup: 'DOCX';
  summary: string;
  capabilities: readonly DocxComplexParseCapability[];
  dependencies: readonly ModuleDependencyRef[];
  configSchema: JsonObject;
}

/** 可注册的模块实例。 */
export interface DocxComplexParseModule {
  definition: DocxComplexParseModuleDefinition;
  handlers: DocxComplexParseHandlers;
}

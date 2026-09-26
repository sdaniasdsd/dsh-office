/**
 * 上游 Profile 模块的「边界类型声明」。
 *
 * 这些是【仅类型】的桩（stub），不是实现。它们精确描述 docx-inspect 从其兄弟模块
 * 消费到的全部接口表面，不多声明一个字段：
 *
 *   - `office-core`     : IR / artifact / warning 等公共原语（本模块输出的目标形态）
 *   - `office-safety`   : `verify` 接口接收的安全策略形状
 *   - `office-files`    : artifact 物化端口（已声明，尚未接线）
 *   - `office-test-kit` : 验证报告需满足的形状契约
 *
 * 为什么用桩而不是真实依赖包？
 *   1. 让本模块可以「单独启动、单独跑 contract test」，无需先构建整个 Profile；
 *   2. 避免复制其他模块的实现（spec 明确禁止复制公开契约的实现）；
 *   3. 正式接入 Profile 时，真实的 `office-core` 等包会覆盖同名模块声明，
 *      本模块代码无需任何改动。
 *
 * 关键约束：`src/` 永远不通过运行时 import 这些模块（由 `verbatimModuleSyntax`
 * 加上 `import type` 双重保证），因此本模块在 Profile 之外也能独立运行与测试。
 *
 * 维护纪律：不要把这些声明扩写成兄弟契约的副本。只有当 docx-inspect 确实要消费
 * 一个新字段时，才在此处补充。
 */

declare module 'office-core' {
  /** 产物在 Profile 内的唯一标识。 */
  export type ArtifactId = string;
  /** 产物的定位符（本地路径、file: URI，或由 office-files 物化的其他 scheme）。 */
  export type ArtifactUri = string;

  /** 指向一个由 Profile 托管的字节产物的引用。 */
  export interface ArtifactRef {
    /** 稳定标识；跨模块传递时必须保持不变。 */
    id: ArtifactId;
    /** 定位符；默认解析器只支持本地路径与 file: URI。 */
    uri: ArtifactUri;
    /** IANA 媒体类型（已知时填写）。 */
    mediaType?: string;
    /** 字节数。 */
    sizeBytes?: number;
    /** 小写十六进制的 SHA-256 摘要。 */
    sha256?: string;
    /** 供人阅读的标签，例如文件名。 */
    label?: string;
    /** 自由标签；值必须是可 JSON 序列化的字符串。 */
    tags?: Record<string, string>;
  }

  /** 告警严重级别。 */
  export type WarningSeverity = 'info' | 'warn' | 'error';

  /** 面向编排层的非致命诊断信息。 */
  export interface Warning {
    /** 稳定、可编程判断的告警码。 */
    code: string;
    /** 人类可读的说明。 */
    message: string;
    severity: WarningSeverity;
    /** 类 JSON 指针的位置，例如 `word/_rels/document.xml.rels`。 */
    path?: string;
    /** 附加上下文；必须可 JSON 序列化。 */
    details?: Record<string, unknown>;
  }

  /** 粗粒度的格式身份。 */
  export type FormatKind =
    | 'docx'
    | 'docm'
    | 'dotx'
    | 'dotm'
    | 'ole'
    | 'rtf'
    | 'zip'
    | 'unknown';

  /** `inspect` 的结果：回答「这是什么产物」。 */
  export interface FormatProfile {
    format: FormatKind;
    /** 主文档部件的媒体类型；无法判定时为 null。 */
    mediaType: string | null;
    /** 依据文件名得到的扩展名；无扩展名时为 null。 */
    extension: string | null;
    /** 0..1 的置信度，由多少个独立信号互相印证推导。 */
    confidence: number;
    container: 'zip' | 'ole' | 'rtf' | 'unknown';
    /** 是否加密；null 表示「无法判断」而非「未加密」。 */
    encrypted: boolean | null;
    /** 命中的探测信号，例如 ['magic:zip', 'content-types:wordprocessingml']。 */
    signatures: string[];
    /** 指标汇总，供策略层/编排层消费。 */
    features: Record<string, boolean>;
    /** 附加元数据；必须可 JSON 序列化。 */
    metadata: Record<string, unknown>;
  }

  /** 包内单个部件（part）的信息。 */
  export interface PackagePart {
    /** 包内路径，例如 `word/document.xml`。 */
    name: string;
    /** 解压后字节数。 */
    sizeBytes: number;
    /** ZIP 中压缩后的字节数。 */
    compressedSize: number;
  }

  /** 包内一条 OPC 关系（relationship）。 */
  export interface PackageRelationship {
    /** 关系的来源部件；包根为 `''`。 */
    sourcePart: string;
    /** 关系 Id，例如 `rId4`。 */
    id: string;
    /** 关系类型 URI。 */
    type: string;
    /** 目标；外部目标的模式见 targetMode。 */
    target: string;
    /** 内部引用还是外部引用。 */
    targetMode: 'Internal' | 'External';
  }

  /** 宏指标：只说明「存在宏」，不判定其内容或风险。 */
  export interface MacroIndicator {
    /** vba = VBA 工程；xlm = Excel 4.0 宏表。 */
    kind: 'vba' | 'xlm';
    /** 承载宏的部件路径。 */
    part: string;
    sizeBytes: number;
  }

  /** 外部引用指标。 */
  export interface ExternalReference {
    sourcePart: string;
    relationshipId: string;
    relationshipType: string;
    /** 外部目标，例如 URL、UNC 路径或 DDE 命令。 */
    target: string;
    /** 分类标签，用于策略做差异化处理。 */
    category: string;
  }

  /** 嵌入对象指标。 */
  export interface EmbeddedObject {
    part: string;
    /** oleEmbedding / package / altChunk / activeX / flash。 */
    kind: string;
    mediaType: string | null;
    sizeBytes: number;
    name: string | null;
  }

  /** `execute` 的结果：与底层引擎无关的归一化结构。 */
  export interface FormatIR {
    profile: FormatProfile;
    parts: PackagePart[];
    relationships: PackageRelationship[];
    indicators: {
      macros: MacroIndicator[];
      externalReferences: ExternalReference[];
      embeddedObjects: EmbeddedObject[];
    };
  }
}

declare module 'office-safety' {
  /**
   * `verify` 接口所评估的安全策略。
   *
   * 所有开关都是可选的：缺省表示「本策略未声明该要求」，对应的检查会被标记为
   * `skip`（跳过）而不是 `pass`（通过）。这样策略作者能明确区分
   * 「我确认允许」与「我根本没表态」。
   */
  export interface SafetyPolicy {
    /** 策略标识，会回填到验证报告中。 */
    id: string;
    /** 是否允许文档携带宏。 */
    allowMacros?: boolean;
    /** 是否允许外部链接（远程模板、超链接等）。 */
    allowExternalLinks?: boolean;
    /** 是否允许嵌入对象。 */
    allowEmbeddedObjects?: boolean;
    /** 是否允许 altChunk（可引入外部内容块）。 */
    allowAltChunks?: boolean;
    /** 是否允许加密产物。 */
    allowEncrypted?: boolean;
    /** 外部关系目标数量的上限。 */
    maxExternalTargets?: number;
    /** 为 true 时，携带宏的文档必须转人工复核。 */
    reviewOnMacros?: boolean;
  }
}

declare module 'office-files' {
  /** 一个本地文件路径，附带可选的清理钩子。 */
  export interface MaterializedArtifact {
    path: string;
    cleanup?: () => Promise<void>;
  }

  /**
   * 把 `ArtifactRef` 物化为可读本地路径的端口。
   *
   * docx-inspect 的默认实现只理解本地路径与 `file:` URI；当 Profile 需要支持
   * 远程/压缩包内的 artifact 时，注入真实实现即可，本模块代码无需改动。
   */
  export interface ArtifactMaterializer {
    materialize(ref: { id: string; uri: string }): Promise<MaterializedArtifact>;
  }
}

declare module 'office-test-kit' {
  /** `VerificationReport` 必须满足的形状契约。 */
  export interface VerificationReportLike {
    /** 是否所有「已声明」的检查都通过。 */
    ok: boolean;
    /** 是否为部分验证（存在跳过项或无法完成的检查）。 */
    partial: boolean;
    checks: unknown[];
    summary: {
      total: number;
      passed: number;
      failed: number;
      skipped: number;
    };
  }
}

// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.

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

/**
 * 上游 Profile 模块的「边界类型声明」。
 *
 * 这些是【仅类型】的桩（stub），不是实现。它们精确描述 docx-parse 从其兄弟模块
 * 消费到的全部接口表面，不多声明一个字段：
 *
 *   - `office-core`     : IR / artifact / warning 等公共原语
 *   - `office-safety`   : `verify` 复用到的安全策略原语（id / allowEncrypted）
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
 * 维护纪律：不要把这些声明扩写成兄弟契约的副本。只有当 docx-parse 确实要消费
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
    /** 类 JSON 指针的位置，例如 `word/document.xml`。 */
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
}

declare module 'office-safety' {
  /**
   * 共享的安全策略原语。
   *
   * docx-parse 的校验聚焦「结构质量」，因此只复用与本模块相关的字段
   * （策略标识与加密处置），安全指标类判定属于 docx-inspect。
   */
  export interface SafetyPolicy {
    /** 策略标识，会回填到验证报告中。 */
    id: string;
    /** 是否允许加密产物。 */
    allowEncrypted?: boolean;
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
   * docx-parse 的默认实现只理解本地路径与 `file:` URI；当 Profile 需要支持
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

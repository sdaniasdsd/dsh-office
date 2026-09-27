/**
 * 映射器 —— 引擎中立探测结果与本模块公开输出之间的【唯一】翻译层。
 *
 * 两项职责，都被刻意放在引擎之外：
 *   1. 裁决结果：哪些失败是致命的、分别对应哪个错误码；
 *   2. 归一化数据：把指标数据整理成 office-core 的 IR 形状，
 *      并保持稳定排序，使输出可 diff、可做快照测试。
 *
 * 引擎无权选择错误码，也无权决定 IR 形状；这些决定都由本文件做出。
 */
import type {
  ArtifactRef,
  EmbeddedObject as IrEmbeddedObject,
  ExternalReference as IrExternalReference,
  FormatIR,
  FormatKind,
  FormatProfile,
  MacroIndicator as IrMacroIndicator,
  PackagePart,
  PackageRelationship,
  Warning,
} from 'office-core';

import type { DocxInspectOptions, ModuleConfig } from './contract';
import {
  hasActiveX,
  hasAltChunks,
  hasDdeFields,
  hasEmbeddedObjects,
  hasExternalReferences,
  hasMacros,
  type ProbeIssue,
  type ProbeResult,
} from './domain/docx-inspect';
import { DocxInspectError } from './errors';

/* -------------------------------------------------------------------------- */
/* 声明式提示（扩展名 / MIME）                                                   */
/* -------------------------------------------------------------------------- */

/** 文档变体 -> 公开的格式种类。 */
const VARIANT_TO_FORMAT: Record<string, FormatKind> = {
  document: 'docx',
  macroEnabled: 'docm',
  template: 'dotx',
  macroEnabledTemplate: 'dotm',
};

/** 文档变体 -> 规范扩展名。 */
const VARIANT_TO_EXTENSION: Record<string, string> = {
  document: 'docx',
  macroEnabled: 'docm',
  template: 'dotx',
  macroEnabledTemplate: 'dotm',
};

/** 本模块认同的 Word 扩展名白名单。 */
const WORD_EXTENSIONS = new Set(['docx', 'docm', 'dotx', 'dotm']);

/** 本模块认同的 Word 媒体类型白名单。 */
const WORD_MEDIA_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  'application/vnd.ms-word.document.macroenabled.12',
  'application/vnd.ms-word.template.macroenabledtemplate.12',
]);

/** 只表示「我不知道类型」的通用媒体类型，不能当作类型冲突。 */
const GENERIC_MEDIA_TYPES = new Set([
  'application/octet-stream',
  'application/zip',
  'application/x-zip-compressed',
  'binary/octet-stream',
]);

/** 一处 声明值 与 探测值 的冲突。 */
export interface HintConflict {
  field: 'declaredExtension' | 'declaredMimeType';
  declared: string;
  detected: string;
}

/** 声明式提示的比对结果。 */
export interface HintEvaluation {
  /** 硬冲突：使产物无法被本模块处理的冲突（对应 FORMAT_MISMATCH）。 */
  hardConflicts: HintConflict[];
  /** 软冲突：值得提示但不阻断（例如用 `docx` 声称了一个宏启用文档）。 */
  softConflicts: HintConflict[];
}

/**
 * 比对调用方给出的提示与实际字节内容。
 *
 * 这是 spec 要求的「类型不匹配」防护，但保持务实：
 * - 通用 Content-Type（octet-stream、zip 等）不算冲突，因为浏览器/邮件网关
 *   经常这样发送，属于「调用方没表态」而非「调用方说错了」；
 * - `application/msword` 只算软冲突：它常被当作 .docx 的近似值使用。
 *
 * @param probe   引擎探测结果
 * @param options 调用方给出的声明式提示
 */
export function evaluateDeclaredHints(
  probe: ProbeResult,
  options: DocxInspectOptions = {},
): HintEvaluation {
  const evaluation: HintEvaluation = { hardConflicts: [], softConflicts: [] };
  const detectedExtension = VARIANT_TO_EXTENSION[probe.documentVariant ?? ''] ?? null;

  // --- 扩展名比对 -------------------------------------------------------- //
  const declaredExtension = options.declaredExtension?.trim().toLowerCase();
  if (declaredExtension !== undefined && declaredExtension !== '') {
    if (!WORD_EXTENSIONS.has(declaredExtension)) {
      // 声称的是完全不相干的扩展名（例如 .pdf / .exe）→ 硬冲突。
      evaluation.hardConflicts.push({
        field: 'declaredExtension',
        declared: declaredExtension,
        detected: detectedExtension ?? probe.extension ?? 'unknown',
      });
    } else if (detectedExtension !== null && declaredExtension !== detectedExtension) {
      // 都是 Word 扩展名，但具体变体不一致（如 docx 声称 vs docm 实际）→ 软冲突。
      evaluation.softConflicts.push({
        field: 'declaredExtension',
        declared: declaredExtension,
        detected: detectedExtension,
      });
    }
  }

  // --- 媒体类型比对 ------------------------------------------------------ //
  const declaredMimeType = options.declaredMimeType?.trim().toLowerCase();
  if (declaredMimeType !== undefined && declaredMimeType !== '') {
    const detectedMime = (probe.mediaType ?? 'unknown').toLowerCase();
    if (GENERIC_MEDIA_TYPES.has(declaredMimeType)) {
      // 调用方未表态，视为无冲突。
    } else if (declaredMimeType === 'application/msword') {
      // 遗留 MIME，仅软冲突。
      evaluation.softConflicts.push({
        field: 'declaredMimeType',
        declared: declaredMimeType,
        detected: detectedMime,
      });
    } else if (!WORD_MEDIA_TYPES.has(declaredMimeType)) {
      evaluation.hardConflicts.push({
        field: 'declaredMimeType',
        declared: declaredMimeType,
        detected: detectedMime,
      });
    }
  }

  return evaluation;
}

/* -------------------------------------------------------------------------- */
/* 结果裁决                                                                     */
/* -------------------------------------------------------------------------- */

/** 引擎错误类型 -> 本模块错误码。 */
const ERROR_KIND_TO_CODE = {
  NOT_FOUND: 'ARTIFACT_NOT_FOUND',
  PERMISSION_DENIED: 'ARTIFACT_NOT_FOUND',
  BAD_ZIP: 'PARSE_FAILED',
  READ_FAILED: 'PARSE_FAILED',
  IO_ERROR: 'PARSE_FAILED',
} as const;

/**
 * 校验探测结果描述的是「本模块可以处理的产物」，否则抛出本模块的错误。
 *
 * 放在任何映射之前调用：这样失败绝不会以「一半构建好的输出」的形式泄漏出去。
 *
 * 检查顺序即为「希望的失败顺序」——越靠前的检查越可能是调用方最想先知道的原因。
 *
 * @throws DocxInspectError 带有恰当的错误码与可序列化 details
 */
export function assertProbeUsable(
  probe: ProbeResult,
  config: ModuleConfig,
  options: DocxInspectOptions = {},
): void {
  // 1. 引擎已判定致命错误 → 按其 kind 翻译成模块错误码。
  if (probe.error !== null) {
    const mapped = ERROR_KIND_TO_CODE[probe.error.kind as keyof typeof ERROR_KIND_TO_CODE];
    throw new DocxInspectError(mapped ?? 'PARSE_FAILED', probe.error.message, {
      details: { engineErrorKind: probe.error.kind },
    });
  }

  // 2. 容器类型超出范围 → 与「格式不符」区分开，语义更准确。
  if (probe.container === 'ole') {
    throw new DocxInspectError(
      'UNSUPPORTED_CONTAINER',
      'OLE/CFB containers (legacy .doc or encrypted OOXML) are outside docx-inspect scope',
      { details: { container: probe.container } },
    );
  }
  if (probe.container === 'rtf') {
    throw new DocxInspectError('UNSUPPORTED_CONTAINER', 'RTF is outside docx-inspect scope', {
      details: { container: probe.container },
    });
  }

  // 3. 连容器都无法识别 → 格式不匹配。
  if (probe.container !== 'zip') {
    throw new DocxInspectError(
      'FORMAT_MISMATCH',
      'Artifact is not a recognizable Office container',
      { details: { container: probe.container } },
    );
  }

  // 4. 命中资源预算：仅在 enforceLimits 开启时视为致命。
  //    关闭时保留结果，让上层按「尽力而为」处理（配合 LIMIT_APPLIED 告警）。
  if (probe.limitHit !== null && config.featureFlags.enforceLimits) {
    throw new DocxInspectError(
      'LIMIT_EXCEEDED',
      `Artifact exceeds the configured budget (${probe.limitHit})`,
      { details: { limit: probe.limitHit, limits: { ...config.limits } } },
    );
  }

  // 5. 是 ZIP，但不是 WordprocessingML → xlsx/pptx 属于其他模块的范围。
  if (probe.documentKind !== 'wordprocessingml') {
    throw new DocxInspectError(
      'FORMAT_MISMATCH',
      'Package is not a WordprocessingML (DOCX) document',
      {
        details: {
          documentKind: probe.documentKind,
          documentVariant: probe.documentVariant,
        },
      },
    );
  }

  // 6. 空包（partCount 为 0）说明不是一个有效的 OPC 包。
  if (probe.partCount === 0) {
    throw new DocxInspectError('PARSE_FAILED', 'Package contains no parts');
  }

  // 7. 声明式提示的硬冲突 → 类型不匹配。
  const hints = evaluateDeclaredHints(probe, options);
  const [conflict] = hints.hardConflicts;
  if (conflict !== undefined) {
    throw new DocxInspectError(
      'FORMAT_MISMATCH',
      `Declared ${conflict.field} does not match the artifact content`,
      {
        details: {
          field: conflict.field,
          declared: conflict.declared,
          detected: conflict.detected,
        },
      },
    );
  }
}

/* -------------------------------------------------------------------------- */
/* 告警                                                                         */
/* -------------------------------------------------------------------------- */

/** 引擎问题码 -> 告警码与严重级别。未登记的码统一降级为 PARTIALLY_INSPECTED。 */
const ISSUE_TO_WARNING: Record<string, { code: string; severity: Warning['severity'] }> = {
  EMPTY_DOCUMENT: { code: 'EMPTY_DOCUMENT', severity: 'info' },
  XML_INVALID: { code: 'PARTIALLY_INSPECTED', severity: 'warn' },
  ENTRY_UNREADABLE: { code: 'PARTIALLY_INSPECTED', severity: 'warn' },
  LIMIT_ENTRY_BYTES: { code: 'LIMIT_APPLIED', severity: 'warn' },
  LIMIT_ARCHIVE_ENTRIES: { code: 'LIMIT_APPLIED', severity: 'warn' },
  LIMIT_REACHED: { code: 'LIMIT_APPLIED', severity: 'warn' },
  ENCRYPTED_PACKAGE: { code: 'ENCRYPTED_ARTIFACT', severity: 'warn' },
};

/** 把单条引擎问题翻译成模块告警。 */
function warningFromIssue(issue: ProbeIssue): Warning {
  const mapped = ISSUE_TO_WARNING[issue.code] ?? {
    code: 'PARTIALLY_INSPECTED' as const,
    severity: 'warn' as const,
  };
  const warning: Warning = {
    code: mapped.code,
    message: issue.message,
    severity: mapped.severity,
    // 保留原始问题码，便于排查「引擎说了什么」。
    details: { issueCode: issue.code },
  };
  if (issue.path !== undefined) {
    warning.path = issue.path;
  }
  return warning;
}

/**
 * 汇总指标与提示类告警，顺序固定。
 *
 * 固定顺序的意义：同一份文档每次调用得到的 warnings 数组完全一致，
 * 这样才能在回归测试中做精确断言。
 */
export function toWarnings(probe: ProbeResult, options: DocxInspectOptions = {}): Warning[] {
  const warnings: Warning[] = [];

  // 引擎上报的问题优先（它们是「事实层」信息）。
  for (const issue of probe.issues) {
    warnings.push(warningFromIssue(issue));
  }

  // XML 后端不是 lxml：结果仍可用，但能力略有差异，需提示。
  if (probe.xmlBackend !== 'lxml') {
    warnings.push({
      code: 'XML_BACKEND_FALLBACK',
      message: `Probe used the "${probe.xmlBackend}" XML backend instead of lxml`,
      severity: 'info',
      details: { xmlBackend: probe.xmlBackend },
    });
  }

  // 以下为「指标类」告警：只陈述存在性，不做风险判定。
  if (hasMacros(probe)) {
    warnings.push({
      code: 'MACROS_PRESENT',
      message: 'Document carries VBA/XLM macro parts',
      severity: 'warn',
      details: {
        count: probe.macroIndicators.length,
        // 排序后再放入：避免引擎输出顺序影响结果。
        parts: probe.macroIndicators.map((indicator) => indicator.part).sort(),
      },
    });
  }

  // 宏语义分析结论（可选增强）。
  // 与 MACROS_PRESENT 的分工：后者陈述「有宏」这一事实，这里给出「宏带恶意行为特征」的判断。
  // 该告警仍属指标层——是否阻断交由 office-safety 决定，模块不自作主张。
  const suspiciousMacros = probe.macroAnalyses.filter((analysis) => analysis.suspicious);
  if (suspiciousMacros.length > 0) {
    warnings.push({
      code: 'MACRO_SUSPICIOUS_CODE',
      message: 'Macro code matches malicious behaviour heuristics (auto-exec plus shell/write)',
      severity: 'warn',
      details: {
        count: suspiciousMacros.length,
        // 全部排序后再放入：避免第三方库的产出顺序影响输出，保证结果可 diff。
        modules: uniqueSorted(suspiciousMacros.map((analysis) => analysis.module)),
        flags: uniqueSorted(suspiciousMacros.map((analysis) => analysis.flags)),
        // 命中关键字即是判定证据，回传给上层便于审计与展示。
        matches: uniqueSorted(suspiciousMacros.flatMap((analysis) => analysis.matches)),
      },
    });
  }

  if (hasExternalReferences(probe)) {
    warnings.push({
      code: 'EXTERNAL_REFERENCES_PRESENT',
      message: 'Document declares external relationship targets',
      severity: 'warn',
      details: {
        count: probe.externalReferences.length,
        categories: uniqueSorted(probe.externalReferences.map((reference) => reference.category)),
      },
    });
  }

  if (hasEmbeddedObjects(probe)) {
    warnings.push({
      code: 'EMBEDDED_OBJECTS_PRESENT',
      message: 'Document embeds OLE/package objects',
      severity: 'warn',
      details: {
        count: probe.embeddedObjects.filter((object) => object.kind !== 'altChunk').length,
      },
    });
  }

  if (hasAltChunks(probe)) {
    warnings.push({
      code: 'ALT_CHUNKS_PRESENT',
      message: 'Document contains altChunk parts',
      severity: 'warn',
      details: {
        count: probe.embeddedObjects.filter((object) => object.kind === 'altChunk').length,
      },
    });
  }

  if (hasActiveX(probe)) {
    warnings.push({
      code: 'ACTIVE_X_PRESENT',
      message: 'Document contains ActiveX controls',
      severity: 'warn',
      details: {
        count: probe.embeddedObjects.filter((object) => object.kind === 'activeX').length,
      },
    });
  }

  if (hasDdeFields(probe)) {
    warnings.push({
      code: 'DDE_FIELDS_PRESENT',
      message: 'Document contains DDE/DDEAUTO fields',
      severity: 'warn',
      details: { count: probe.ddeFields.length },
    });
  }

  // 软冲突（声明与探测不一致但不阻断）也作为告警上报。
  const hints = evaluateDeclaredHints(probe, options);
  for (const conflict of hints.softConflicts) {
    warnings.push({
      code: 'DECLARED_HINT_MISMATCH',
      message: `Declared ${conflict.field} "${conflict.declared}" disagrees with detected "${conflict.detected}"`,
      severity: 'warn',
      details: { ...conflict },
    });
  }

  return warnings;
}

/* -------------------------------------------------------------------------- */
/* FormatProfile                                                                */
/* -------------------------------------------------------------------------- */

/**
 * 计算 0..1 的置信度。
 *
 * 权重设计原则：越难「偶然撞上」的信号权重越高。
 * `magic:zip` 很容易被任意 zip 满足（低权重），
 * `wordprocessingml` 的主部件声明则几乎不可能是巧合（高权重）。
 */
function computeConfidence(probe: ProbeResult): number {
  let score = 0;
  if (probe.container === 'zip') score += 0.3;
  if (probe.documentKind === 'wordprocessingml') score += 0.4;
  if (probe.documentVariant !== null) score += 0.2;
  const detectedExtension = VARIANT_TO_EXTENSION[probe.documentVariant ?? ''];
  // 扩展名与内容互证：额外加分。
  if (detectedExtension !== undefined && probe.extension === detectedExtension) score += 0.1;
  // 保留两位小数，避免浮点尾数污染快照测试。
  return Math.min(1, Math.round(score * 100) / 100);
}

/** 公开的「这是什么产物？」答案。 */
export function toFormatProfile(
  probe: ProbeResult,
  options: DocxInspectOptions = {},
): FormatProfile {
  const format = VARIANT_TO_FORMAT[probe.documentVariant ?? ''] ?? 'unknown';

  // 记录命中的探测信号，便于排查「为什么判定成这个格式」。
  const signatures: string[] = [];
  if (probe.container === 'zip') signatures.push('magic:zip');
  if (probe.documentKind === 'wordprocessingml') signatures.push('content-types:wordprocessingml');
  if (probe.documentVariant !== null) signatures.push(`variant:${probe.documentVariant}`);
  const detectedExtension = VARIANT_TO_EXTENSION[probe.documentVariant ?? ''];
  if (detectedExtension !== undefined && probe.extension === detectedExtension) {
    signatures.push(`extension:${detectedExtension}`);
  }

  const hints = evaluateDeclaredHints(probe, options);

  return {
    format,
    mediaType: probe.mediaType,
    extension: probe.extension,
    confidence: computeConfidence(probe),
    container: probe.container,
    encrypted: probe.encrypted,
    signatures,
    // 指标汇总：供策略层/编排层直接消费，无需自行遍历明细。
    features: {
      hasMacros: hasMacros(probe),
      hasExternalLinks: hasExternalReferences(probe),
      hasEmbeddedObjects: hasEmbeddedObjects(probe),
      hasAltChunks: hasAltChunks(probe),
      hasActiveX: hasActiveX(probe),
      hasDdeFields: hasDdeFields(probe),
      declaredHintsCompatible: hints.hardConflicts.length === 0,
    },
    // 元数据：只放规模与版本信息，不放文档内容。
    metadata: {
      probeVersion: probe.probeVersion,
      xmlBackend: probe.xmlBackend,
      documentVariant: probe.documentVariant,
      partCount: probe.partCount,
      relationshipCount: probe.relationships.length,
      macroIndicatorCount: probe.macroIndicators.length,
      externalReferenceCount: probe.externalReferences.length,
      embeddedObjectCount: probe.embeddedObjects.length,
      ddeFieldCount: probe.ddeFields.length,
      limitHit: probe.limitHit,
    },
  };
}

/* -------------------------------------------------------------------------- */
/* FormatIR                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * 归一化为与引擎无关的包结构，供下游模块消费。
 *
 * 所有数组都按稳定键排序：这样 IR 具备确定性，
 * 下游可以安全地对它做哈希、diff 或快照。
 */
export function toFormatIR(probe: ProbeResult, options: DocxInspectOptions = {}): FormatIR {
  const parts: PackagePart[] = [...probe.parts]
    .sort((left, right) => compareStrings(left.name, right.name))
    .map((part) => ({
      name: part.name,
      sizeBytes: part.sizeBytes,
      compressedSize: part.compressedSize,
    }));

  const relationships: PackageRelationship[] = [...probe.relationships]
    .sort(
      (left, right) =>
        compareStrings(left.sourcePart, right.sourcePart) || compareStrings(left.id, right.id),
    )
    .map((relationship) => ({ ...relationship }));

  const macros: IrMacroIndicator[] = [...probe.macroIndicators]
    .sort(
      (left, right) =>
        compareStrings(left.part, right.part) || compareStrings(left.kind, right.kind),
    )
    .map((indicator) => ({ ...indicator }));

  const externalReferences: IrExternalReference[] = [...probe.externalReferences]
    .sort(
      (left, right) =>
        compareStrings(left.target, right.target) ||
        compareStrings(left.sourcePart, right.sourcePart) ||
        compareStrings(left.relationshipId, right.relationshipId),
    )
    .map((reference) => ({ ...reference }));

  const embeddedObjects: IrEmbeddedObject[] = [...probe.embeddedObjects]
    .sort(
      (left, right) =>
        compareStrings(left.part, right.part) || compareStrings(left.kind, right.kind),
    )
    .map((object) => ({ ...object }));

  return {
    profile: toFormatProfile(probe, options),
    parts,
    relationships,
    indicators: { macros, externalReferences, embeddedObjects },
  };
}

/* -------------------------------------------------------------------------- */
/* artifact 引用补全                                                            */
/* -------------------------------------------------------------------------- */

/**
 * 返回调用方引用的副本，并用探测所得信息补全它。
 *
 * 关键约束：模块【绝不】伪造新的 id —— artifact 跨模块流转时必须保持身份不变，
 * 否则 Profile 无法把它们关联起来。
 */
export function refineArtifactRef(artifactRef: ArtifactRef, probe: ProbeResult): ArtifactRef {
  const refined: ArtifactRef = {
    ...artifactRef,
    sizeBytes: probe.sizeBytes,
  };
  if (probe.sha256 !== '') {
    refined.sha256 = probe.sha256;
  }
  if (probe.mediaType !== null) {
    refined.mediaType = probe.mediaType;
  }
  // 只在调用方没有提供 label 时才用「document.<ext>」兜底，不覆盖调用方的命名。
  if (probe.extension !== null) {
    refined.label = artifactRef.label ?? `document.${probe.extension}`;
  }
  return refined;
}

/** 字符串比较器：显式定义是为了得到与 locale 无关的稳定顺序。 */
function compareStrings(left: string, right: string): number {
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

/** 去重并排序。 */
function uniqueSorted(values: string[]): string[] {
  return [...new Set(values)].sort();
}

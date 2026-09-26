/**
 * 映射器 —— 引擎中立解析结果与本模块公开输出之间的【唯一】翻译层。
 *
 * 三项职责，都被刻意放在引擎之外：
 *   1. 裁决结果：哪些失败是致命的、分别对应哪个错误码；
 *   2. 派生视图：由 blocks 推导 outline / counts（单一数据源，不信任引擎自报）；
 *   3. 归一化顺序：对「顺序无序」的集合稳定排序，使 IR 可 diff、可做快照测试。
 *
 * 注意：本模块【不做双 IR】——`ParseResult`（引擎载荷经信任边界校验后的领域对象）
 * 就是 IR 的数据部分；本文件只做「判可用 → 补 profile → 派生 → 排序」，
 * 不重塑结构。引擎无权选择错误码，也无权决定 IR 形状。
 */
import type { ArtifactRef, FormatKind, FormatProfile, Warning } from 'office-core';

import type { DocxParseOptions, ModuleConfig } from './contract';
import {
  buildOutline,
  computeCounts,
  hasVisibleContent,
  type Block,
  type CommentRecord,
  type DocxParseIR,
  type FootnoteRecord,
  type ParseIssue,
  type ParseResult,
  type RelationshipRecord,
  type StyleRecord,
} from './domain/docx-parse';
import { DocxParseError } from './errors';

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
 * - 通用 Content-Type（octet-stream、zip 等）不算冲突；
 * - `application/msword` 只算软冲突：它常被当作 .docx 的近似值使用。
 */
export function evaluateDeclaredHints(
  result: ParseResult,
  options: DocxParseOptions = {},
): HintEvaluation {
  const evaluation: HintEvaluation = { hardConflicts: [], softConflicts: [] };
  const detectedExtension = VARIANT_TO_EXTENSION[result.documentVariant ?? ''] ?? null;

  const declaredExtension = options.declaredExtension?.trim().toLowerCase();
  if (declaredExtension !== undefined && declaredExtension !== '') {
    if (!WORD_EXTENSIONS.has(declaredExtension)) {
      evaluation.hardConflicts.push({
        field: 'declaredExtension',
        declared: declaredExtension,
        detected: detectedExtension ?? result.extension ?? 'unknown',
      });
    } else if (detectedExtension !== null && declaredExtension !== detectedExtension) {
      evaluation.softConflicts.push({
        field: 'declaredExtension',
        declared: declaredExtension,
        detected: detectedExtension,
      });
    }
  }

  const declaredMimeType = options.declaredMimeType?.trim().toLowerCase();
  if (declaredMimeType !== undefined && declaredMimeType !== '') {
    const detectedMime = (result.mediaType ?? 'unknown').toLowerCase();
    if (GENERIC_MEDIA_TYPES.has(declaredMimeType)) {
      // 调用方未表态，视为无冲突。
    } else if (declaredMimeType === 'application/msword') {
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
 * 校验解析结果描述的是「本模块可以处理的产物」，否则抛出本模块的错误。
 *
 * 放在任何映射之前调用：这样失败绝不会以「一半构建好的输出」的形式泄漏出去。
 * 检查顺序即为「希望的失败顺序」——越靠前的检查越可能是调用方最想先知道的原因。
 *
 * @throws DocxParseError 带有恰当的错误码与可序列化 details
 */
export function assertParseUsable(
  result: ParseResult,
  config: ModuleConfig,
  options: DocxParseOptions = {},
): void {
  // 1. 引擎已判定致命错误 → 按其 kind 翻译成模块错误码。
  if (result.error !== null) {
    const mapped = ERROR_KIND_TO_CODE[result.error.kind as keyof typeof ERROR_KIND_TO_CODE];
    throw new DocxParseError(mapped ?? 'PARSE_FAILED', result.error.message, {
      details: { engineErrorKind: result.error.kind },
    });
  }

  // 2. 容器类型超出范围 → 与「格式不符」区分开，语义更准确。
  if (result.container === 'ole') {
    throw new DocxParseError(
      'UNSUPPORTED_CONTAINER',
      'OLE/CFB containers (legacy .doc or encrypted OOXML) are outside docx-parse scope',
      { details: { container: result.container } },
    );
  }
  if (result.container === 'rtf') {
    throw new DocxParseError('UNSUPPORTED_CONTAINER', 'RTF is outside docx-parse scope', {
      details: { container: result.container },
    });
  }

  // 3. 连容器都无法识别 → 格式不匹配。
  if (result.container !== 'zip') {
    throw new DocxParseError(
      'FORMAT_MISMATCH',
      'Artifact is not a recognizable Office container',
      { details: { container: result.container } },
    );
  }

  // 4. 命中资源预算：仅在 enforceLimits 开启时视为致命。
  //    关闭时保留结果，让上层按「尽力而为」处理（配合 LIMIT_APPLIED 告警）。
  if (result.limitHit !== null && config.featureFlags.enforceLimits) {
    throw new DocxParseError(
      'LIMIT_EXCEEDED',
      `Artifact exceeds the configured budget (${result.limitHit})`,
      { details: { limit: result.limitHit, limits: { ...config.limits } } },
    );
  }

  // 5. 是 ZIP，但不是 WordprocessingML → xlsx/pptx 属于其他模块的范围。
  if (result.documentKind !== 'wordprocessingml') {
    throw new DocxParseError(
      'FORMAT_MISMATCH',
      'Package is not a WordprocessingML (DOCX) document',
      {
        details: {
          documentKind: result.documentKind,
          documentVariant: result.documentVariant,
        },
      },
    );
  }

  // 6. 声明式提示的硬冲突 → 类型不匹配。
  const hints = evaluateDeclaredHints(result, options);
  const [conflict] = hints.hardConflicts;
  if (conflict !== undefined) {
    throw new DocxParseError(
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

/**
 * 引擎问题码 -> 告警码与严重级别。
 *
 * 这是一张【适配表】：引擎侧的问题码是开放词汇（引擎可独立演进，也可能被
 * 调用方替换），模块侧的告警码是封闭词汇。未登记的码统一降级为
 * `PARTIALLY_PARSED` 并保留原始码在 `details.issueCode` —— 因此「引擎发了
 * 未知的码」不会丢信息，新增引擎能力也不必先改这里。
 *
 * 反过来不成立：本表【不应】为引擎从不发出的码登记映射，否则等于给一个
 * 不存在的产出路径造出告警，且没有任何测试会失败。新增映射的前提是引擎
 * 确实会发对应的码。
 *
 * 导出仅为让测试断言「每个映射键都有引擎产出路径」（见 contract.test.ts）；
 * 它不构成公开契约。
 */
export const ISSUE_TO_WARNING: Record<string, { code: string; severity: Warning['severity'] }> = {
  EMPTY_DOCUMENT: { code: 'EMPTY_DOCUMENT', severity: 'info' },
  XML_INVALID: { code: 'PARTIALLY_PARSED', severity: 'warn' },
  ENTRY_UNREADABLE: { code: 'PARTIALLY_PARSED', severity: 'warn' },
  LIMIT_REACHED: { code: 'LIMIT_APPLIED', severity: 'warn' },
  STYLES_MISSING: { code: 'STYLES_MISSING', severity: 'warn' },
  STYLE_NOT_FOUND: { code: 'BROKEN_STYLE_REFERENCE', severity: 'warn' },
  DANGLING_RELATIONSHIP: { code: 'DANGLING_RELATIONSHIP', severity: 'warn' },
  UNKNOWN_ELEMENT: { code: 'UNKNOWN_BLOCK_SKIPPED', severity: 'info' },
};

/** 把单条引擎问题翻译成模块告警。 */
function warningFromIssue(issue: ParseIssue): Warning {
  const mapped = ISSUE_TO_WARNING[issue.code] ?? {
    code: 'PARTIALLY_PARSED' as const,
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
export function toWarnings(
  result: ParseResult,
  options: DocxParseOptions = {},
  config?: ModuleConfig,
): Warning[] {
  const warnings: Warning[] = [];

  // 引擎上报的问题优先（它们是「事实层」信息）。
  for (const issue of result.issues) {
    warnings.push(warningFromIssue(issue));
  }

  // 预算类事实以 `limitHit` 为准（引擎问题列表可能被截断，但该字段不会）。
  if (result.limitHit !== null && !warnings.some((warning) => warning.code === 'LIMIT_APPLIED')) {
    warnings.push({
      code: 'LIMIT_APPLIED',
      message: `Artifact exceeded the "${result.limitHit}" budget`,
      severity: 'warn',
      details: { limit: result.limitHit },
    });
  }

  // XML 后端不是 lxml：结果仍可用，但能力略有差异，需提示。
  if (result.xmlBackend !== 'lxml') {
    warnings.push({
      code: 'XML_BACKEND_FALLBACK',
      message: `Parser used the "${result.xmlBackend}" XML backend instead of lxml`,
      severity: 'info',
      details: { xmlBackend: result.xmlBackend },
    });
  }

  // 以下为「指标类」告警：只陈述结构事实，不做风险判定。
  const counts = computeCounts(result);
  const parseHeadings = config?.featureFlags.parseHeadings ?? true;
  if (parseHeadings && counts.headings === 0 && hasVisibleContent(result.blocks)) {
    warnings.push({
      code: 'NO_HEADINGS',
      message: 'Document contains no headings',
      severity: 'info',
      details: { paragraphs: counts.paragraphs, tables: counts.tables },
    });
  }

  if (counts.comments > 0) {
    warnings.push({
      code: 'COMMENTS_PRESENT',
      message: 'Document contains comments',
      severity: 'info',
      details: { count: counts.comments },
    });
  }

  if (counts.footnotes + counts.endnotes > 0) {
    warnings.push({
      code: 'FOOTNOTES_PRESENT',
      message: 'Document contains footnotes or endnotes',
      severity: 'info',
      details: { footnotes: counts.footnotes, endnotes: counts.endnotes },
    });
  }

  // 软冲突（声明与探测不一致但不阻断）也作为告警上报。
  const hints = evaluateDeclaredHints(result, options);
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
function computeConfidence(result: ParseResult): number {
  let score = 0;
  if (result.container === 'zip') score += 0.3;
  if (result.documentKind === 'wordprocessingml') score += 0.4;
  if (result.documentVariant !== null) score += 0.2;
  const detectedExtension = VARIANT_TO_EXTENSION[result.documentVariant ?? ''];
  if (detectedExtension !== undefined && result.extension === detectedExtension) score += 0.1;
  // 保留两位小数，避免浮点尾数污染快照测试。
  return Math.min(1, Math.round(score * 100) / 100);
}

/** 公开的「这是什么产物？」答案。 */
export function toFormatProfile(
  result: ParseResult,
  options: DocxParseOptions = {},
): FormatProfile {
  const format = VARIANT_TO_FORMAT[result.documentVariant ?? ''] ?? 'unknown';
  const counts = computeCounts(result);

  const signatures: string[] = [];
  if (result.container === 'zip') signatures.push('magic:zip');
  if (result.documentKind === 'wordprocessingml') signatures.push('content-types:wordprocessingml');
  if (result.documentVariant !== null) signatures.push(`variant:${result.documentVariant}`);
  const detectedExtension = VARIANT_TO_EXTENSION[result.documentVariant ?? ''];
  if (detectedExtension !== undefined && result.extension === detectedExtension) {
    signatures.push(`extension:${detectedExtension}`);
  }

  const hints = evaluateDeclaredHints(result, options);

  return {
    format,
    mediaType: result.mediaType,
    extension: result.extension,
    confidence: computeConfidence(result),
    container: result.container,
    encrypted: result.encrypted,
    signatures,
    features: {
      hasHeadings: counts.headings > 0,
      hasTables: counts.tables > 0,
      hasStyles: counts.styles > 0,
      hasRelationships: counts.relationships > 0,
      hasComments: counts.comments > 0,
      hasFootnotes: counts.footnotes + counts.endnotes > 0,
      emptyDocument: !hasVisibleContent(result.blocks),
      declaredHintsCompatible: hints.hardConflicts.length === 0,
    },
    metadata: {
      parseVersion: result.parseVersion,
      xmlBackend: result.xmlBackend,
      documentVariant: result.documentVariant,
      blockCount: counts.blocks,
      paragraphCount: counts.paragraphs,
      tableCount: counts.tables,
      headingCount: counts.headings,
      styleCount: counts.styles,
      relationshipCount: counts.relationships,
      commentCount: counts.comments,
      footnoteCount: counts.footnotes,
      endnoteCount: counts.endnotes,
      limitHit: result.limitHit,
    },
  };
}

/* -------------------------------------------------------------------------- */
/* DocxParseIR                                                                  */
/* -------------------------------------------------------------------------- */

/** 克隆一个正文块，避免公开 IR 与内部结果共享引用。 */
function cloneBlock(block: Block): Block {
  if (block.kind === 'paragraph') {
    return { kind: 'paragraph', paragraph: { ...block.paragraph, runs: block.paragraph.runs.map((run) => ({ ...run })) } };
  }
  return { kind: 'table', table: { ...block.table, cells: block.table.cells.map((cell) => ({ ...cell })) } };
}

/**
 * 归一化为与引擎无关、且顺序确定的解析 IR。
 *
 * 排序策略：
 *   - `blocks` / `outline` 保持文档顺序（顺序本身是信息，不可排序）；
 *   - `styles` / `relationships` / `comments` / `footnotes` 按稳定键排序，
 *     这样下游可以安全地对 IR 做哈希、diff 或快照。
 */
export function toDocxParseIR(result: ParseResult, options: DocxParseOptions = {}): DocxParseIR {
  const styles: StyleRecord[] = [...result.styles]
    .sort((left, right) => compareStrings(left.styleId, right.styleId))
    .map((style) => ({ ...style }));

  const relationships: RelationshipRecord[] = [...result.relationships]
    .sort(
      (left, right) =>
        compareStrings(left.sourcePart, right.sourcePart) || compareStrings(left.id, right.id),
    )
    .map((relationship) => ({ ...relationship }));

  const comments: CommentRecord[] = [...result.comments]
    .sort((left, right) => compareNotes(left.id, right.id))
    .map((comment) => ({ ...comment }));

  const footnotes: FootnoteRecord[] = [...result.footnotes]
    .sort(
      (left, right) =>
        compareStrings(left.kind, right.kind) || compareNotes(left.id, right.id),
    )
    .map((note) => ({ ...note }));

  const blocks = result.blocks.map(cloneBlock);

  return {
    profile: toFormatProfile(result, options),
    metadata: { ...result.metadata },
    outline: buildOutline(blocks),
    blocks,
    styles,
    relationships,
    comments,
    footnotes,
    counts: computeCounts(result),
  };
}

/* -------------------------------------------------------------------------- */
/* artifact 引用补全                                                            */
/* -------------------------------------------------------------------------- */

/**
 * 返回调用方引用的副本，并用解析所得信息补全它。
 *
 * 关键约束：模块【绝不】伪造新的 id —— artifact 跨模块流转时必须保持身份不变，
 * 否则 Profile 无法把它们关联起来。
 */
export function refineArtifactRef(artifactRef: ArtifactRef, result: ParseResult): ArtifactRef {
  const refined: ArtifactRef = {
    ...artifactRef,
    sizeBytes: result.sizeBytes,
  };
  if (result.sha256 !== '') {
    refined.sha256 = result.sha256;
  }
  if (result.mediaType !== null) {
    refined.mediaType = result.mediaType;
  }
  // 只在调用方没有提供 label 时才用「document.<ext>」兜底，不覆盖调用方的命名。
  if (result.extension !== null) {
    refined.label = artifactRef.label ?? `document.${result.extension}`;
  }
  return refined;
}

/** 字符串比较器：显式定义是为了得到与 locale 无关的稳定顺序。 */
function compareStrings(left: string, right: string): number {
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

/** 数值优先的 id 比较器（脚注 / 批注 id 多为数字字符串）。 */
function compareNotes(left: string, right: string): number {
  const leftNum = Number.parseInt(left, 10);
  const rightNum = Number.parseInt(right, 10);
  const leftIsNum = Number.isFinite(leftNum) && String(leftNum) === left;
  const rightIsNum = Number.isFinite(rightNum) && String(rightNum) === right;
  if (leftIsNum && rightIsNum && leftNum !== rightNum) {
    return leftNum < rightNum ? -1 : 1;
  }
  return compareStrings(left, right);
}

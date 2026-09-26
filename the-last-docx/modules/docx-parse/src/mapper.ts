/**
 * 映射器 —— 引擎中立观察与双 IR / office-core IR 之间的【唯一】翻译层。
 *
 * 三项职责，都被刻意放在引擎之外：
 *   1. 裁决结果：哪些失败是致命的、分别对应哪个错误码；
 *   2. 构建双 IR：把原始观察归一化成「语义视图 + 物理视图 + 对应表」；
 *   3. 映射 IR：把双 IR 装进 office-core 的 `FormatIR` 形状。
 *
 * 引擎无权选择错误码，也无权决定 IR 形状与节点身份。
 */
import type {
  ArtifactRef,
  EmbeddedObject as IrEmbeddedObject,
  ExternalReference as IrExternalReference,
  FormatKind,
  FormatProfile,
  MacroIndicator as IrMacroIndicator,
  PackagePart,
  PackageRelationship,
  Warning,
} from 'office-core';

import type {
  DocxDualIR,
  DocxParseIR,
  DocxParseOptions,
  ModuleConfig,
  NodeAnchor,
  PhysicalDocument,
  PhysicalNode,
  SemanticAnnotation,
  SemanticBlock,
  SemanticCell,
  SemanticParagraph,
  SemanticRow,
  SemanticTable,
  SourceMap,
} from './contract';
import { DUAL_VIEW_COUNT, SOURCE_MAP_SCHEME } from './contract';
import type {
  ParseIssue,
  ParseResult,
  RawAnnotation,
  RawBlock,
  RawParagraph,
} from './domain/docx-parse';
import { DocxParseError } from './errors';
import {
  buildAnchor,
  computeSemanticId,
  isAnchorWeak,
  normalizeText,
  SourceMapBuilder,
} from './sourcemap';

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
 * 与 docx-inspect 保持同一套判定口径：通用 Content-Type 不算冲突；
 * `application/msword` 只算软冲突。两个模块对「类型是否匹配」的结论必须一致。
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

  // 2. 容器类型超出范围。
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
    throw new DocxParseError('FORMAT_MISMATCH', 'Artifact is not a recognizable Office container', {
      details: { container: result.container },
    });
  }

  // 4. 命中资源预算：仅在 enforceLimits 开启时视为致命。
  if (result.limitHit !== null && config.featureFlags.enforceLimits) {
    throw new DocxParseError(
      'LIMIT_EXCEEDED',
      `Artifact exceeds the configured budget (${result.limitHit})`,
      { details: { limit: result.limitHit, limits: { ...config.limits } } },
    );
  }

  // 5. 是 ZIP，但不是 WordprocessingML → xlsx/pptx 属于其他模块的范围。
  if (result.documentKind !== 'wordprocessingml') {
    throw new DocxParseError('FORMAT_MISMATCH', 'Package is not a WordprocessingML (DOCX) document', {
      details: {
        documentKind: result.documentKind,
        documentVariant: result.documentVariant,
      },
    });
  }

  // 6. 空包（partCount 为 0）说明不是一个有效的 OPC 包。
  if (result.partCount === 0) {
    throw new DocxParseError('PARSE_FAILED', 'Package contains no parts');
  }

  // 7. 声明式提示的硬冲突 → 类型不匹配。
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

/** 引擎问题码 -> 告警码与严重级别。未登记的码统一降级为 PARTIALLY_PARSED。 */
const ISSUE_TO_WARNING: Record<string, { code: string; severity: Warning['severity'] }> = {
  EMPTY_DOCUMENT: { code: 'EMPTY_DOCUMENT', severity: 'info' },
  STYLES_UNREADABLE: { code: 'STYLE_FALLBACK', severity: 'warn' },
  STYLES_MISSING: { code: 'STYLE_FALLBACK', severity: 'warn' },
  UNSUPPORTED_ELEMENT: { code: 'UNSUPPORTED_CONTENT', severity: 'info' },
  XML_INVALID: { code: 'PARTIALLY_PARSED', severity: 'warn' },
  ENTRY_UNREADABLE: { code: 'PARTIALLY_PARSED', severity: 'warn' },
  LIMIT_ENTRY_BYTES: { code: 'LIMIT_APPLIED', severity: 'warn' },
  LIMIT_ARCHIVE_ENTRIES: { code: 'LIMIT_APPLIED', severity: 'warn' },
  LIMIT_REACHED: { code: 'LIMIT_APPLIED', severity: 'warn' },
  ENCRYPTED_PACKAGE: { code: 'ENCRYPTED_ARTIFACT', severity: 'warn' },
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
 * 汇总告警，顺序固定。
 *
 * 固定顺序的意义：同一份文档每次调用得到的 warnings 数组完全一致，
 * 这样才能在回归测试中做精确断言。
 */
export function toWarnings(
  result: ParseResult,
  content: DocxDualIR,
  config: ModuleConfig,
  options: DocxParseOptions = {},
): Warning[] {
  const warnings: Warning[] = [];

  // 引擎上报的问题优先（它们是「事实层」信息）。
  for (const issue of result.issues) {
    warnings.push(warningFromIssue(issue));
  }

  // 正文为空：告警而非错误——它能被正常解析，只是没有内容。
  if (content.semantic.blocks.length === 0) {
    warnings.push({
      code: 'EMPTY_DOCUMENT',
      message: 'Document body contains no paragraphs or tables.',
      severity: 'info',
      details: { blockCount: 0 },
    });
  }

  // 标题层级靠推断而来：级别可能不准，需要上层注意。
  const inferred = content.semantic.blocks.filter(
    (block) => block.kind === 'heading' && block.levelSource === 'inferred',
  );
  if (inferred.length > 0) {
    warnings.push({
      code: 'HEADING_LEVEL_INFERRED',
      message: 'Some headings have no explicit outline level; levels were inferred.',
      severity: 'warn',
      details: { count: inferred.length, ids: inferred.map((block) => block.id).sort() },
    });
  }

  // 锚点冗余度不足：抗漂移能力退化，回写定位可能失准。
  if (config.featureFlags.resolveAnchors) {
    const weak = content.sourceMap.entries.filter((entry) => isAnchorWeak(entry.anchor));
    if (weak.length > 0) {
      warnings.push({
        code: 'ANCHOR_INCOMPLETE',
        message: 'Some nodes carry only one anchor selector, weakening edit-time relocation.',
        severity: 'warn',
        details: { count: weak.length, ids: weak.map((entry) => entry.id).sort() },
      });
    }
  }

  // 孤儿注释：在正文里找不到锚点。
  const orphanComments = content.semantic.annotations.filter(
    (annotation) => annotation.kind === 'comment' && annotation.anchorBlockId === null,
  );
  if (orphanComments.length > 0) {
    warnings.push({
      code: 'COMMENT_ORPHANED',
      message: 'Some comments have no resolvable anchor in the document body.',
      severity: 'warn',
      details: { count: orphanComments.length, references: orphanComments.map((a) => a.reference).sort() },
    });
  }

  const orphanNotes = content.semantic.annotations.filter(
    (annotation) => annotation.kind !== 'comment' && annotation.anchorBlockId === null,
  );
  if (orphanNotes.length > 0) {
    warnings.push({
      code: 'NOTE_ORPHANED',
      message: 'Some footnotes/endnotes have no resolvable reference in the document body.',
      severity: 'warn',
      details: { count: orphanNotes.length, references: orphanNotes.map((a) => a.reference).sort() },
    });
  }

  // 预算命中但被宽免。
  if (result.limitHit !== null && !config.featureFlags.enforceLimits) {
    warnings.push({
      code: 'LIMIT_APPLIED',
      message: `Artifact exceeded the "${result.limitHit}" budget; parsing continued leniently.`,
      severity: 'warn',
      details: { limit: result.limitHit },
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

/** 计算 0..1 的置信度。 */
function computeConfidence(result: ParseResult): number {
  let score = 0;
  if (result.container === 'zip') score += 0.3;
  if (result.documentKind === 'wordprocessingml') score += 0.4;
  if (result.documentVariant !== null) score += 0.2;
  const detectedExtension = VARIANT_TO_EXTENSION[result.documentVariant ?? ''];
  if (detectedExtension !== undefined && result.extension === detectedExtension) score += 0.1;
  return Math.min(1, Math.round(score * 100) / 100);
}

/** 公开的「这是什么产物？」答案。 */
export function toFormatProfile(
  result: ParseResult,
  content: DocxDualIR,
  options: DocxParseOptions = {},
): FormatProfile {
  const format = VARIANT_TO_FORMAT[result.documentVariant ?? ''] ?? 'unknown';

  const signatures: string[] = [];
  if (result.container === 'zip') signatures.push('magic:zip');
  if (result.documentKind === 'wordprocessingml') signatures.push('content-types:wordprocessingml');
  if (result.documentVariant !== null) signatures.push(`variant:${result.documentVariant}`);
  const detectedExtension = VARIANT_TO_EXTENSION[result.documentVariant ?? ''];
  if (detectedExtension !== undefined && result.extension === detectedExtension) {
    signatures.push(`extension:${detectedExtension}`);
  }

  const hints = evaluateDeclaredHints(result, options);
  const hasHeadings = content.semantic.blocks.some((block) => block.kind === 'heading');

  return {
    format,
    mediaType: result.mediaType,
    extension: result.extension,
    confidence: computeConfidence(result),
    container: result.container,
    encrypted: result.encrypted,
    signatures,
    features: {
      hasHeadings,
      hasTables: content.semantic.blocks.some((block) => block.kind === 'table'),
      hasComments: content.semantic.annotations.some((a) => a.kind === 'comment'),
      hasFootnotes: content.semantic.annotations.some((a) => a.kind === 'footnote'),
      hasEndnotes: content.semantic.annotations.some((a) => a.kind === 'endnote'),
      declaredHintsCompatible: hints.hardConflicts.length === 0,
    },
    metadata: {
      parseVersion: result.parseVersion,
      xmlBackend: result.xmlBackend,
      mainPart: result.mainPart,
      blockCount: content.semantic.blocks.length,
      annotationCount: content.semantic.annotations.length,
      sourceMapEntries: content.sourceMap.entries.length,
    },
  };
}

/* -------------------------------------------------------------------------- */
/* 双 IR 构建                                                                   */
/* -------------------------------------------------------------------------- */

/** 标题级别的裁决结果。 */
interface HeadingDecision {
  level: number;
  source: 'outlineLevel' | 'styleName' | 'inferred';
}

/** 把 1..9 之外的级别夹回合法区间。 */
function clampHeadingLevel(level: number): number {
  if (!Number.isFinite(level)) return 1;
  if (level < 1) return 1;
  if (level > 9) return 9;
  return Math.floor(level);
}

/**
 * 裁决一个段落是否为标题、以及级别是几级。
 *
 * 优先级刻意设计为 outlineLvl > 样式名 > 标题样式默认值：
 *   - outlineLvl 是 Word 内部的「大纲级别」，最权威且与界面语言无关；
 *   - 样式名（"Heading 2"）会被本地化（中文版是「标题 2」），只能作为次选；
 *   - 只用标题样式但拿不到级别时，退化为 1 级，并如实标记为 inferred。
 *
 * 这正应对「标题识别依据」这个经典两难：不能只看样式名，也不能假设 outlineLvl 一定存在。
 */
function decideHeading(raw: RawParagraph, config: ModuleConfig): HeadingDecision | null {
  if (!config.featureFlags.parseHeadings) return null;
  if (raw.outlineLevel !== null) {
    return { level: clampHeadingLevel(raw.outlineLevel + 1), source: 'outlineLevel' };
  }
  if (raw.styleNameLevel !== null) {
    return { level: clampHeadingLevel(raw.styleNameLevel), source: 'styleName' };
  }
  if (raw.headingStyle) {
    return { level: 1, source: 'inferred' };
  }
  return null;
}

/** 由一个原始段落构造语义段落。 */
function toSemanticParagraph(raw: RawParagraph): SemanticParagraph {
  const anchor = buildParagraphAnchor(raw);
  return {
    id: computeSemanticId(anchor),
    kind: 'paragraph',
    anchor,
    text: normalizeText(raw.text),
    styleId: raw.styleId,
    styleName: raw.styleName,
    list: raw.list,
  };
}

/** 构造段落锚点。 */
function buildParagraphAnchor(raw: RawParagraph): NodeAnchor {
  return buildAnchor({
    kind: 'p',
    part: raw.part,
    structuralPath: raw.path,
    ordinal: raw.ordinal,
    paraId: raw.paraId,
    text: raw.text,
  });
}

/** 构造块（段落/表格）锚点。 */
function buildBlockAnchor(block: RawBlock): NodeAnchor {
  return buildAnchor({
    kind: block.kind === 'table' ? 'tbl' : 'p',
    part: block.part,
    structuralPath: block.path,
    ordinal: block.ordinal,
    paraId: block.kind === 'paragraph' ? block.paraId : null,
    text: block.text,
  });
}

/** 构造注释锚点（位于 comments.xml / footnotes.xml / endnotes.xml 内）。 */
function buildAnnotationAnchor(annotation: RawAnnotation): NodeAnchor {
  return buildAnchor({
    kind: annotation.kind,
    part: annotation.part,
    structuralPath: annotation.path,
    ordinal: annotation.ordinal,
    paraId: null,
    text: annotation.text,
  });
}

/** 一个语义块 + 其对应的物理指针，供 source map 登记。 */
interface BlockProjection {
  block: SemanticBlock;
  pointers: string[];
}

/**
 * 裁决注释锚定到哪个正文块。
 *
 * 注释可能跨多个段落（锚点列表非空），取第一个能在正文里找到的块。
 * 一个都找不到 → 孤儿注释（anchorBlockId 为 null），由告警如实上报。
 */
function resolveAnnotationAnchor(
  annotation: RawAnnotation,
  blockIdByPointer: Map<string, string>,
): string | null {
  for (const pointer of annotation.anchorPointers) {
    const blockId = blockIdByPointer.get(pointer);
    if (blockId !== undefined) {
      return blockId;
    }
  }
  return null;
}

/**
 * 把一个原始块投影成语义块。
 *
 * 表格的单元格段落同样获得独立的语义 id 与锚点——这样「改表格里某一格的一段话」
 * 可以精确落到那个 `<w:p>` 上，而不是被迫重写整张表（这正是「最小 diff」的前提）。
 */
function projectBlock(
  raw: RawBlock,
  config: ModuleConfig,
  builder: SourceMapBuilder,
  physical: PhysicalNode[],
): BlockProjection {
  const anchor = buildBlockAnchor(raw);
  const id = computeSemanticId(anchor);

  if (raw.kind === 'paragraph') {
    const heading = decideHeading(raw, config);
    const block: SemanticBlock =
      heading === null
        ? toSemanticParagraph(raw)
        : {
            id,
            kind: 'heading',
            anchor,
            level: heading.level,
            levelSource: heading.source,
            text: normalizeText(raw.text),
            styleId: raw.styleId,
            styleName: raw.styleName,
          };
    builder.add({
      id: block.id,
      kind: block.kind,
      pointer: raw.pointer,
      pointers: [raw.pointer],
      anchor: block.anchor,
    });
    physical.push(toPhysicalNode(block.anchor, raw.pointer, raw.text));
    return { block, pointers: [raw.pointer] };
  }

  // 表格：为表格本身与每个单元格段落分别登记。
  const rows: SemanticRow[] = raw.rows.map((row) => ({
    cells: row.cells.map((cell) => {
      const paragraphs = cell.paragraphs.map((paragraph) => {
        const semanticParagraph = toSemanticParagraph(paragraph);
        builder.add({
          id: semanticParagraph.id,
          kind: semanticParagraph.kind,
          pointer: paragraph.pointer,
          pointers: [paragraph.pointer],
          anchor: semanticParagraph.anchor,
        });
        physical.push(toPhysicalNode(semanticParagraph.anchor, paragraph.pointer, paragraph.text));
        return semanticParagraph;
      });
      const semanticCell: SemanticCell = {
        text: normalizeText(cell.paragraphs.map((p) => p.text).join('\n')),
        gridSpan: cell.gridSpan,
        vMerge: cell.vMerge,
        paragraphs,
      };
      return semanticCell;
    }),
  }));

  const table: SemanticTable = {
    id,
    kind: 'table',
    anchor,
    gridColumns: raw.gridColumns,
    rows,
  };
  builder.add({
    id: table.id,
    kind: table.kind,
    pointer: raw.pointer,
    pointers: [raw.pointer],
    anchor: table.anchor,
  });
  physical.push(toPhysicalNode(table.anchor, raw.pointer, raw.text));
  return { block: table, pointers: [raw.pointer] };
}

/** 由锚点构造物理视图节点。 */
function toPhysicalNode(anchor: NodeAnchor, pointer: string, text: string): PhysicalNode {
  return {
    pointer,
    part: anchor.part,
    kind: anchor.kind,
    ordinal: anchor.ordinal,
    paraId: anchor.paraId,
    digest: anchor.digest,
    textLength: normalizeText(text).length,
  };
}

/**
 * 构建双 IR。
 *
 * 步骤：投影正文块（同时登记 + 建物理视图）→ 投影注释 → 组装对应表。
 * 所有输出都按稳定键排序，保证同一份文档得到完全一致的 IR。
 */
export function toContentIR(result: ParseResult, config: ModuleConfig): DocxDualIR {
  const builder = new SourceMapBuilder();
  const physicalNodes: PhysicalNode[] = [];
  const blocks: SemanticBlock[] = [];
  const blockIdByPointer = new Map<string, string>();

  for (const raw of result.blocks) {
    if (raw.kind === 'table' && !config.featureFlags.parseTables) {
      // 关闭表格解析：仍保留为段落级占位会误导上层，因此跳过并交由告警解释。
      continue;
    }
    const projection = projectBlock(raw, config, builder, physicalNodes);
    blocks.push(projection.block);
    blockIdByPointer.set(projection.pointers[0] ?? '', projection.block.id);
  }

  const annotations: SemanticAnnotation[] = [];
  if (config.featureFlags.parseComments || config.featureFlags.parseFootnotes) {
    for (const raw of result.annotations) {
      if (raw.kind === 'comment' && !config.featureFlags.parseComments) continue;
      if (raw.kind !== 'comment' && !config.featureFlags.parseFootnotes) continue;

      const anchor = buildAnnotationAnchor(raw);
      const annotation: SemanticAnnotation = {
        id: computeSemanticId(anchor),
        kind: raw.kind,
        reference: raw.reference,
        text: normalizeText(raw.text),
        author: raw.author,
        date: raw.date,
        anchorBlockId: resolveAnnotationAnchor(raw, blockIdByPointer),
        anchor,
      };
      builder.add({
        id: annotation.id,
        kind: annotation.kind,
        pointer: raw.pointer,
        pointers: [raw.pointer],
        anchor,
      });
      physicalNodes.push(toPhysicalNode(anchor, raw.pointer, raw.text));
      annotations.push(annotation);
    }
  }

  const sourceMap: SourceMap = config.featureFlags.resolveAnchors
    ? builder.build()
    : { scheme: SOURCE_MAP_SCHEME, entries: [], bySemanticId: {}, byPointer: {}, byFingerprint: {} };

  const physical: PhysicalDocument = {
    view: 'physical',
    mainPart: result.mainPart,
    // 物理节点按指针排序，保证视图稳定。
    nodes: [...physicalNodes].sort((left, right) => compareStrings(left.pointer, right.pointer)),
  };

  return {
    viewCount: DUAL_VIEW_COUNT,
    scheme: SOURCE_MAP_SCHEME,
    semantic: { view: 'semantic', blocks, annotations },
    physical,
    sourceMap,
  };
}

/* -------------------------------------------------------------------------- */
/* DocxParseIR                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * 归一化为 office-core 的 IR，并在 `content` 字段承载双 IR。
 *
 * 所有数组都按稳定键排序：这样 IR 具备确定性，下游可以安全地做哈希、diff 或快照。
 */
export function toDocxParseIR(
  result: ParseResult,
  config: ModuleConfig,
  options: DocxParseOptions = {},
): DocxParseIR {
  const content = toContentIR(result, config);

  const parts: PackagePart[] = [...result.parts]
    .sort((left, right) => compareStrings(left.name, right.name))
    .map((part) => ({
      name: part.name,
      sizeBytes: part.sizeBytes,
      compressedSize: part.compressedSize,
    }));

  const relationships: PackageRelationship[] = [...result.relationships]
    .sort(
      (left, right) =>
        compareStrings(left.sourcePart, right.sourcePart) || compareStrings(left.id, right.id),
    )
    .map((relationship) => ({ ...relationship }));

  // 解析模块不产出安全指标：这几项恒为空，语义由 docx-inspect / office-safety 负责。
  const macros: IrMacroIndicator[] = [];
  const externalReferences: IrExternalReference[] = [];
  const embeddedObjects: IrEmbeddedObject[] = [];

  return {
    profile: toFormatProfile(result, content, options),
    parts,
    relationships,
    indicators: { macros, externalReferences, embeddedObjects },
    content,
  };
}

/* -------------------------------------------------------------------------- */
/* artifact 引用补全                                                            */
/* -------------------------------------------------------------------------- */

/**
 * 返回调用方引用的副本，并用解析所得信息补全它。
 *
 * 关键约束：模块【绝不】伪造新的 id —— artifact 跨模块流转时必须保持身份不变。
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

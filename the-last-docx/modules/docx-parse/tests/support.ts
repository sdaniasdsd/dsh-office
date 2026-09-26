/**
 * 测试支撑：构造「引擎无关」的解析结果与测试替身。
 *
 * 为什么需要它：本模块的绝大部分行为（结果裁决、双 IR 构建、对应表、验证）
 * 都发生在引擎【之后】。用一份手工构造的 `ParseResult` 就能精确覆盖这些路径，
 * 而不必先生成 fixture、再请出 Python。真正的 Python 集成只在回归测试里出现一次。
 */
import { fileURLToPath } from 'node:url';

import type { ArtifactRef } from 'office-core';
import type { SafetyPolicy } from 'office-safety';

import type { DocxEngine } from '../src/engine/adapter';
import type {
  ParseResult,
  RawAnnotation,
  RawBlock,
  RawParagraph,
  RawTable,
} from '../src/domain/docx-parse';
import type { DocxDualIR } from '../src/contract';
import type { RevisionIntent } from '../src/revision/contract';
import { toContentIR } from '../src/mapper';
import { resolveConfig } from '../src/config';

export const PROJECT_ROOT = fileURLToPath(new URL('..', import.meta.url));
export const FIXTURES_DIR = fileURLToPath(new URL('../fixtures/_generated/', import.meta.url));

/** 拼出 fixture 的绝对路径。 */
export function fixture(name: string): string {
  return `${FIXTURES_DIR}${name}`;
}

/** 构造一个 artifact 引用（测试里只关心 uri）。 */
export function artifactRef(uri: string, extra: Partial<ArtifactRef> = {}): ArtifactRef {
  return { id: 'artifact-1', uri, ...extra };
}

/** 构造一个段落块。 */
export function para(
  text: string,
  index: number,
  extra: Partial<Omit<RawParagraph, 'kind'>> = {},
): RawBlock {
  const path = `/w:document/w:body/w:p[${index + 1}]`;
  const paragraph: RawParagraph = {
    pointer: `word/document.xml!${path}`,
    part: 'word/document.xml',
    path,
    ordinal: index,
    paraId: null,
    text,
    styleId: null,
    styleName: null,
    outlineLevel: null,
    styleNameLevel: null,
    headingStyle: false,
    list: null,
    commentRefs: [],
    footnoteRefs: [],
    endnoteRefs: [],
    ...extra,
  };
  return { kind: 'paragraph', ...paragraph };
}

/** 用「文本 -> 一个自动段落」的规则快速铺一段正文。 */
export function paras(texts: string[]): RawBlock[] {
  return texts.map((text, index) => para(text, index));
}

/** 构造一个「文本 -> 一个单元格」的表格块。 */
export function table(rows: string[][], index: number, extra: Partial<Omit<RawTable, 'kind'>> = {}): RawBlock {
  const path = `/w:document/w:body/w:tbl[${index + 1}]`;
  const built: RawTable = {
    pointer: `word/document.xml!${path}`,
    part: 'word/document.xml',
    path,
    ordinal: index,
    text: rows.map((row) => row.join('')).join('\n'),
    gridColumns: rows[0]?.length ?? 0,
    rows: rows.map((row) => ({
      cells: row.map((text) => ({
        gridSpan: 1,
        vMerge: null,
        paragraphs: [
          {
            pointer: `word/document.xml!${path}/w:tr[1]/w:tc[1]/w:p[1]`,
            part: 'word/document.xml',
            path: `${path}/w:tr[1]/w:tc[1]/w:p[1]`,
            ordinal: 0,
            paraId: null,
            text,
            styleId: null,
            styleName: null,
            outlineLevel: null,
            styleNameLevel: null,
            headingStyle: false,
            list: null,
            commentRefs: [],
            footnoteRefs: [],
            endnoteRefs: [],
          },
        ],
      })),
    })),
    ...extra,
  };
  return { kind: 'table', ...built };
}

/** 构造一条注释。 */
export function annotation(
  kind: RawAnnotation['kind'],
  reference: string,
  anchorPointers: string[],
  extra: Partial<RawAnnotation> = {},
): RawAnnotation {
  const part = kind === 'comment' ? 'word/comments.xml' : `word/${kind}s.xml`;
  const wrapper = kind === 'comment' ? 'comments' : `${kind}s`;
  return {
    kind,
    reference,
    text: `${kind} ${reference}`,
    author: kind === 'comment' ? 'Tester' : null,
    date: kind === 'comment' ? '2024-01-02T03:04:05Z' : null,
    pointer: `${part}!/w:${wrapper}/w:${kind}[1]`,
    part,
    path: `/w:${wrapper}/w:${kind}[1]`,
    ordinal: 0,
    anchorPointers,
    ...extra,
  };
}

/** 构造一份完整、可用的 `ParseResult`，只覆盖测试关心的字段。 */
export function parseResult(overrides: Partial<ParseResult> = {}): ParseResult {
  return {
    parseVersion: 1,
    xmlBackend: 'lxml',
    container: 'zip',
    extension: 'docx',
    sizeBytes: 4096,
    sha256: 'a'.repeat(64),
    encrypted: false,
    documentKind: 'wordprocessingml',
    documentVariant: 'document',
    mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    partCount: 3,
    parts: [
      { name: 'word/document.xml', sizeBytes: 1024, compressedSize: 512 },
      { name: 'word/styles.xml', sizeBytes: 512, compressedSize: 256 },
      { name: '[Content_Types].xml', sizeBytes: 256, compressedSize: 128 },
    ],
    relationships: [],
    mainPart: 'word/document.xml',
    blocks: [],
    annotations: [],
    limitHit: null,
    issues: [],
    error: null,
    ...overrides,
  };
}

/** 测试替身引擎：原样返回预置结果，彻底绕开进程外解析。 */
export class FakeEngine implements DocxEngine {
  readonly name = 'fake-engine';
  parseCalls = 0;

  constructor(private readonly result: ParseResult) {}

  async parse(): Promise<ParseResult> {
    this.parseCalls += 1;
    return this.result;
  }

  async dispose(): Promise<void> {
    // 无状态替身，无需清理。
  }
}

/** 一个「什么都不拒绝」的策略：仅用于验证接口形状，不产生 policy 相关失败。 */
export function permissivePolicy(overrides: Partial<SafetyPolicy> = {}): SafetyPolicy {
  return {
    id: 'test-policy',
    allowMacros: false,
    allowExternalLinks: false,
    allowEncrypted: false,
    ...overrides,
  };
}

/* -------------------------------------------------------------------------- */
/* 双 IR 与修订层的测试支撑                                                      */
/* -------------------------------------------------------------------------- */

/**
 * 用「原始块」直接造出一份真实的双 IR。
 *
 * 为什么走真实 mapper 而不是手搓一个假 IR：语义 id、锚点、对应表全是
 * 从原始观察推导出来的。手搓会绕过被推导的部分，测出来的结论就不作数了。
 */
export function buildIR(
  blocks: RawBlock[],
  overrides: Partial<ParseResult> = {},
): DocxDualIR {
  return toContentIR(parseResult({ blocks, ...overrides }), resolveConfig());
}

/** 造一条意图，字段默认值只是为了让测试聚焦在关心的那一两个字段上。 */
export function intent(overrides: Partial<RevisionIntent> = {}): RevisionIntent {
  return {
    id: 'intent-1',
    requestId: 'req-1',
    source: 'user',
    stage: 'revision',
    kind: 'content',
    text: '把这段改短一点',
    targetIds: [],
    ...overrides,
  };
}

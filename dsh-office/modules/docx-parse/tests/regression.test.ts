/**
 * 回归测试 —— 双 IR 的核心不变量。
 *
 * 这是本模块最有价值的一组测试：回写场景没有「权威裁判」，因此我们退而求其次，
 * 用两条【可判定的不变量】守住质量底线：
 *   1. 对应表自洽：语义节点与物理节点严丝合缝，没有孤儿、没有抢占；
 *   2. 身份确定性：同一份输入重算 id 必须逐字节相同。
 * 再加一条工程纪律：同一份文档两次调用必须产出完全一致的 IR（可快照、可 diff）。
 */
import { describe, expect, it } from 'vitest';

import {
  DUAL_VIEW_COUNT,
  SOURCE_MAP_SCHEME,
  SourceMapBuilder,
  anchorSelectorCount,
  buildAnchor,
  computeNodeDigest,
  computeSemanticId,
  createDocxParseModule,
  isAnchorWeak,
  normalizeText,
  runVerification,
} from '../src/index';
import type {
  DocxDualIR,
  ExecuteInput,
  InspectInput,
  VerifyInput,
  VerifyOutput,
} from '../src/index';
import { toContentIR } from '../src/mapper';
import { resolveConfig } from '../src/config';
import { artifactRef, FakeEngine, fixture, parseResult, permissivePolicy } from './support';

const module = createDocxParseModule();

/** 一次 execute 调用。 */
async function executeFixture(name: string) {
  const input: ExecuteInput = {
    artifactRef: artifactRef(fixture(name)),
    operation: 'execute',
    requestId: `regression-execute-${name}`,
  };
  return module.handlers.execute(input);
}

/** 一次 verify 调用。 */
async function verifyFixture(
  name: string,
  policy: VerifyInput['policy'],
): Promise<VerifyOutput> {
  const input: VerifyInput = {
    artifactRef: artifactRef(fixture(name)),
    operation: 'verify',
    requestId: `regression-verify-${name}`,
    policy,
  };
  return module.handlers.verify(input);
}

/** 在双 IR 中找到第一个标题块。 */
function firstHeading(content: DocxDualIR) {
  const block = content.semantic.blocks.find((entry) => entry.kind === 'heading');
  if (block === undefined || block.kind !== 'heading') {
    throw new Error('expected a heading block');
  }
  return block;
}

describe('regression: dual IR structure', () => {
  it('produces exactly two views plus a source map', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;
    expect(content.viewCount).toBe(DUAL_VIEW_COUNT);
    expect(content.scheme).toBe(SOURCE_MAP_SCHEME);
    expect(content.semantic.view).toBe('semantic');
    expect(content.physical.view).toBe('physical');
    expect(content.physical.mainPart).toBe('word/document.xml');
  });

  it('projects blocks in document order', async () => {
    const output = await executeFixture('rich.docx');
    const kinds = output.result.ir.content.semantic.blocks.map((block) => block.kind);
    expect(kinds).toEqual([
      'heading',
      'paragraph',
      'paragraph',
      'paragraph',
      'table',
      'heading',
      'paragraph',
    ]);
  });

  it('resolves heading levels: explicit outlineLvl wins, basedOn chain is honoured', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;
    const headings = content.semantic.blocks.filter((block) => block.kind === 'heading');
    expect(headings).toHaveLength(2);

    const [chapter, section] = headings;
    // Heading1 声明了 outlineLvl=0 -> 级别 1，来源为可验证的 outlineLevel。
    expect(chapter?.level).toBe(1);
    expect(chapter?.levelSource).toBe('outlineLevel');
    // Heading3 自身无 outlineLvl，经 basedOn 继承到 2 -> 级别 3。
    expect(section?.level).toBe(3);
    expect(section?.levelSource).toBe('outlineLevel');
  });

  it('infers heading level from the style name when no outline level exists', async () => {
    const output = await executeFixture('heading-by-name.docx');
    const heading = firstHeading(output.result.ir.content);
    expect(heading.level).toBe(2);
    expect(heading.levelSource).toBe('styleName');
    expect(output.warnings.map((warning) => warning.code)).not.toContain('HEADING_LEVEL_INFERRED');
  });

  it('keeps paraId as the primary identity when present', async () => {
    const output = await executeFixture('rich.docx');
    const heading = firstHeading(output.result.ir.content);
    expect(heading.anchor.paraId).toBe('1A2B3C4D');
    // 以 paraId 为主键算出的 id 必须可由锚点复现。
    expect(computeSemanticId(heading.anchor)).toBe(heading.id);
  });
});

describe('regression: table projection does not collapse cells', () => {
  it('keeps rows, cells and per-cell paragraphs individually addressable', async () => {
    const output = await executeFixture('rich.docx');
    const table = output.result.ir.content.semantic.blocks.find((block) => block.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') throw new Error('expected a table block');

    expect(table.rows).toHaveLength(2);
    expect(table.gridColumns).toBe(2);
    const cellTexts = table.rows.flatMap((row) =>
      row.cells.map((cell) => cell.paragraphs.map((paragraph) => paragraph.text).join('')),
    );
    expect(cellTexts).toEqual(['A1', 'B1', 'A2', 'B2']);

    // 每个单元格段落都要有【自己】的语义 id 与对应表记录——
    // 这正是「只改某一格」能产生最小 diff 的前提。
    const content = output.result.ir.content;
    for (const row of table.rows) {
      for (const cell of row.cells) {
        for (const paragraph of cell.paragraphs) {
          expect(content.sourceMap.bySemanticId[paragraph.id]).toBeDefined();
        }
      }
    }
  });
});

describe('regression: annotations stay attached to the right block', () => {
  it('anchors a comment and a footnote to their referencing paragraph', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;
    const blocks = content.semantic.blocks;

    const comment = content.semantic.annotations.find((entry) => entry.kind === 'comment');
    const footnote = content.semantic.annotations.find((entry) => entry.kind === 'footnote');
    expect(comment).toBeDefined();
    expect(footnote).toBeDefined();

    // 批注锚定在第 3 个块（Commented paragraph），脚注锚定在第 4 个块。
    expect(comment?.anchorBlockId).toBe(blocks[2]?.id);
    expect(footnote?.anchorBlockId).toBe(blocks[3]?.id);
    // 分隔符脚注（id 0 / -1）不得被当成真实脚注。
    expect(content.semantic.annotations.filter((entry) => entry.kind === 'footnote')).toHaveLength(1);
  });
});

describe('regression: source map is self-consistent and reversible', () => {
  it('maps every semantic node to exactly one physical anchor', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;

    for (const entry of content.sourceMap.entries) {
      // 正向：id -> 指针。
      expect(content.sourceMap.bySemanticId[entry.id]).toContain(entry.pointer);
      // 反向：指针 -> id。
      expect(content.sourceMap.byPointer[entry.pointer]).toBe(entry.id);
      // 内容寻址：指纹 -> id。
      expect(content.sourceMap.byFingerprint[entry.anchor.digest]).toContain(entry.id);
    }

    // 7 个正文块 + 4 个单元格段落 + 2 条注释 = 13 个语义节点。
    expect(content.sourceMap.entries).toHaveLength(13);
    // 物理视图登记的是「OOXML 里真实存在的元素」，因此同样包含
    // comments.xml / footnotes.xml 中的注释元素（它们也是可写入的节点）。
    expect(content.physical.nodes).toHaveLength(13);
    expect(content.physical.nodes.filter((node) => node.part === 'word/document.xml')).toHaveLength(11);
  });

  it('carries at least two independent anchor selectors for non-empty nodes', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;
    for (const entry of content.sourceMap.entries) {
      expect(anchorSelectorCount(entry.anchor)).toBeGreaterThanOrEqual(2);
      expect(isAnchorWeak(entry.anchor)).toBe(false);
    }
    expect(output.warnings.map((warning) => warning.code)).not.toContain('ANCHOR_INCOMPLETE');
  });
});

describe('regression: determinism', () => {
  it('produces a byte-identical IR across repeated calls', async () => {
    const first = await executeFixture('rich.docx');
    const second = await executeFixture('rich.docx');
    expect(JSON.stringify(second.result.ir)).toBe(JSON.stringify(first.result.ir));
  });

  it('recomputes every stored id from its anchor', async () => {
    const output = await executeFixture('rich.docx');
    const content = output.result.ir.content;
    for (const entry of content.sourceMap.entries) {
      expect(computeSemanticId(entry.anchor)).toBe(entry.id);
    }
  });

  it('keeps the source map identical when the same document is parsed twice', async () => {
    const first = await executeFixture('rich.docx');
    const second = await executeFixture('rich.docx');
    expect(second.result.ir.content.sourceMap).toEqual(first.result.ir.content.sourceMap);
  });
});

describe('regression: identity layer unit invariants', () => {
  it('prefers paraId over the structural path for identity', () => {
    const base = {
      kind: 'p',
      part: 'word/document.xml',
      structuralPath: '/w:document/w:body/w:p[1]',
      ordinal: 0,
      text: 'same text',
    };
    const withoutId = buildAnchor({ ...base, paraId: null });
    const withId = buildAnchor({ ...base, paraId: 'ABCD1234' });
    // 文本相同、位置相同，仅原生 id 不同 —— id 必须随之不同。
    expect(computeSemanticId(withoutId)).not.toBe(computeSemanticId(withId));
  });

  it('derives the fingerprint from kind and normalised text', () => {
    expect(computeNodeDigest('p', '  Hello   world ')).toBe(computeNodeDigest('p', 'Hello world'));
    // 种类不同则指纹不同（避免段落与表格因同一段文字而混同）。
    expect(computeNodeDigest('p', 'x')).not.toBe(computeNodeDigest('tbl', 'x'));
  });

  it('normalises whitespace deterministically', () => {
    expect(normalizeText('  a\n\t b  ')).toBe('a b');
  });

  it('does not count the content-derived digest as an extra selector', () => {
    const anchor = buildAnchor({
      kind: 'p',
      part: 'word/document.xml',
      structuralPath: '/w:document/w:body/w:p[1]',
      ordinal: 0,
      paraId: null,
      text: '',
    });
    // 空段落且无 paraId：只剩位置选择器，冗余度为 1（弱锚点）。
    expect(anchorSelectorCount(anchor)).toBe(1);
    expect(isAnchorWeak(anchor)).toBe(true);
  });

  it('sorts and de-duplicates the source map indexes', () => {
    const builder = new SourceMapBuilder();
    const anchor = buildAnchor({
      kind: 'p',
      part: 'word/document.xml',
      structuralPath: '/w:document/w:body/w:p[1]',
      ordinal: 0,
      paraId: null,
      text: 'x',
    });
    builder.add({ id: 'b', kind: 'p', pointer: 'p2', pointers: ['p2', 'p2'], anchor });
    builder.add({ id: 'a', kind: 'p', pointer: 'p1', pointers: ['p1'], anchor });
    const map = builder.build();
    expect(map.entries.map((entry) => entry.id)).toEqual(['a', 'b']);
    expect(map.bySemanticId['b']).toEqual(['p2']);
  });
});

describe('regression: verification report', () => {
  it('passes on a clean document when the policy declares no budgets', async () => {
    const output = await verifyFixture('rich.docx', permissivePolicy());
    expect(output.result.ok).toBe(true);
    expect(output.result.summary.failed).toBe(0);
    // 未声明预算 -> skip，因此报告是「部分验证」。
    expect(output.result.partial).toBe(true);
    expect(output.result.checks.every((check) => check.status !== 'fail')).toBe(true);
  });

  it('marks nothing as skipped when every requirement is declared', async () => {
    const output = await verifyFixture(
      'rich.docx',
      permissivePolicy({ maxBlocks: 1000, maxTableCells: 1000 }),
    );
    expect(output.result.ok).toBe(true);
    expect(output.result.partial).toBe(false);
  });

  it('fails and warns when the document exceeds a declared budget', async () => {
    const output = await verifyFixture('rich.docx', permissivePolicy({ maxBlocks: 2 }));
    expect(output.result.ok).toBe(false);
    const blocks = output.result.checks.find((check) => check.id === 'policy.blockBudget');
    expect(blocks?.status).toBe('fail');
    expect(output.warnings.map((warning) => warning.code)).toContain('VERIFICATION_FAILED');
  });

  it('fails the encryption requirement when the artifact is encrypted', async () => {
    const encrypted = createDocxParseModule({
      engine: new FakeEngine(parseResult({ encrypted: true, blocks: [] })),
    });
    const input: VerifyInput = {
      artifactRef: artifactRef('fake.docx'),
      operation: 'verify',
      requestId: 'regression-encrypted',
      policy: permissivePolicy(),
    };
    const output = await encrypted.handlers.verify(input);
    const encryption = output.result.checks.find((check) => check.id === 'policy.encryption');
    expect(encryption?.status).toBe('fail');
    expect(output.result.ok).toBe(false);
  });

  it('skips anchor checks when resolveAnchors is disabled', async () => {
    // 直接对 mapper 产物做验证：关闭锚点解析后应当全为 skip/pass，不应误报失败。
    const result = parseResult({ blocks: [] });
    const config = resolveConfig({ featureFlags: { resolveAnchors: false } });
    const content = toContentIR(result, config);
    const report = runVerification(result, content, permissivePolicy(), config);
    const sourceMap = report.checks.find((check) => check.id === 'sourcemap.integrity');
    const determinism = report.checks.find((check) => check.id === 'identity.deterministic');
    expect(sourceMap?.status).toBe('skip');
    expect(determinism?.status).toBe('skip');
    expect(report.ok).toBe(true);
  });

  it('attaches a self-check report to inspect when asked', async () => {
    const input: InspectInput = {
      artifactRef: artifactRef(fixture('clean.docx')),
      operation: 'inspect',
      requestId: 'regression-self-check',
      options: { verifyAfterInspect: true },
    };
    const output = await module.handlers.inspect(input);
    expect(output.verification?.ok).toBe(true);
  });
});

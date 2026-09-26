/**
 * 边界与失败场景测试 —— 用【真实 Python 引擎】跑 fixture 样本。
 *
 * 覆盖 spec 完成标准第三条：至少覆盖正常、边界、损坏或拒绝场景。
 * 若当前机器没有 Python，整组用例会跳过（契约测试仍会照常执行）。
 */
import { describe, expect, it } from 'vitest';

import { makeModule, artifactRef, pythonAvailable, expectJsonRoundTrip } from './support';
import type { DocxParseModule } from '../src/index';

const hasPython = pythonAvailable();

/** 取一个样本的执行输出。 */
async function execute(module: DocxParseModule, name: string, options: Record<string, unknown> = {}) {
  return module.handlers.execute({
    artifactRef: artifactRef(name),
    operation: 'execute',
    requestId: `req-${name}`,
    ...(Object.keys(options).length > 0 ? { options } : {}),
  });
}

/** 取一个样本的检查输出。 */
async function inspect(module: DocxParseModule, name: string, extra: Record<string, unknown> = {}) {
  return module.handlers.inspect({
    artifactRef: artifactRef(name),
    operation: 'inspect',
    requestId: `req-${name}`,
    ...extra,
  });
}

/** 取一个样本的验证报告。 */
async function verify(
  module: DocxParseModule,
  name: string,
  policy: Record<string, unknown>,
) {
  return module.handlers.verify({
    artifactRef: artifactRef(name),
    operation: 'verify',
    requestId: `req-verify-${name}`,
    policy: { id: 'edge', ...policy },
  });
}

const warningCodes = (warnings: { code: string }[]): string[] =>
  warnings.map((warning) => warning.code);

describe.skipIf(!hasPython)('edge: normal documents', () => {
  it('parses a clean document into a complete IR', async () => {
    const module = makeModule();
    const output = await execute(module, 'clean.docx');
    const { ir, artifact } = output.result;

    expect(ir.profile.format).toBe('docx');
    expect(ir.metadata.title).toBe('Clean Fixture');
    expect(artifact.sha256).toHaveLength(64);
    expect(artifact.sizeBytes).toBeGreaterThan(0);

    // 正文顺序即文档顺序。
    expect(ir.blocks.map((block) => block.kind)).toEqual([
      'paragraph',
      'paragraph',
      'paragraph',
      'table',
    ]);
    expect(ir.blocks[0]?.kind === 'paragraph' && ir.blocks[0].paragraph.text).toBe('Clean Fixture');
    expect(ir.blocks[0]?.kind === 'paragraph' && ir.blocks[0].paragraph.headingLevel).toBe(1);
    // 命名样式会被解析成 styleName。
    expect(ir.blocks[1]?.kind === 'paragraph' && ir.blocks[1].paragraph.styleName).toBe('Body Text');

    expect(ir.counts).toMatchObject({ blocks: 4, paragraphs: 3, tables: 1, headings: 1, styles: 7 });
    expect(ir.outline).toEqual([{ level: 1, text: 'Clean Fixture', paragraphIndex: 0 }]);
    // 表格 2x2 展开为 4 个单元格。
    const tableBlock = ir.blocks[3];
    expect(tableBlock?.kind === 'table' && tableBlock.table.cells).toHaveLength(4);
    expectJsonRoundTrip(output);
  });

  it('emits no warning-severity warnings for a clean document', async () => {
    const output = await inspect(makeModule(), 'clean.docx');
    const severe = output.warnings.filter((warning) => warning.severity !== 'info');
    expect(severe).toEqual([]);
  });

  it('does not mistake Word range markers for unparsed content', async () => {
    // clean.docx 刻意带有真实 Word 形状的块级标记（w:sectPr / w:proofErr / 书签）。
    // 它们不含文本，必须被静默忽略——否则每一份真实文档都会被误报为「有未解析内容」。
    // 注意：这个断言不能用 severity 过滤代替，UNKNOWN_BLOCK_SKIPPED 是 info 级。
    const output = await execute(makeModule(), 'clean.docx');
    expect(warningCodes(output.warnings)).not.toContain('UNKNOWN_BLOCK_SKIPPED');
  });

  it('reports body content it could not parse instead of dropping it silently', async () => {
    const output = await execute(makeModule(), 'unknown-elements.docx');
    // 公式段落不被解析：调用方必须能知道「抽取到的文本 != 文档全部文本」。
    expect(warningCodes(output.warnings)).toContain('UNKNOWN_BLOCK_SKIPPED');
    // 跳过的是公式，前后两个段落仍按文档顺序完整保留。
    expect(output.result.ir.blocks.map((block) => block.kind)).toEqual([
      'paragraph',
      'paragraph',
    ]);
  });

  it('derives a multi-level outline from heading styles', async () => {
    const output = await execute(makeModule(), 'headings.docx');
    expect(output.result.ir.outline.map((entry) => entry.level)).toEqual([1, 2, 3]);
    expect(output.result.ir.outline.map((entry) => entry.text)).toEqual([
      'Chapter One',
      'Section 1.1',
      'Subsection 1.1.1',
    ]);
  });

  it('keeps merged table geometry without unbounded recursion', async () => {
    const output = await execute(makeModule(), 'tables.docx');
    const block = output.result.ir.blocks.find((entry) => entry.kind === 'table');
    expect(block?.kind).toBe('table');
    if (block?.kind !== 'table') return;
    expect(block.table.rows).toBe(3);
    expect(block.table.columns).toBe(2);
    // 第一行是一个跨两列的单元格。
    expect(block.table.cells[0]).toMatchObject({ row: 0, column: 0, columnSpan: 2 });
    expect(block.table.cells).toHaveLength(5);
  });

  it('recognizes the macro-enabled variant', async () => {
    const output = await inspect(makeModule(), 'macro.docm');
    expect(output.result.format).toBe('docm');
    expect(output.result.extension).toBe('docm');
  });
});

describe.skipIf(!hasPython)('edge: structural edge cases', () => {
  it('reports an empty document as a warning, not an error', async () => {
    const output = await inspect(makeModule(), 'empty.docx');
    expect(output.result.features['emptyDocument']).toBe(true);
    expect(warningCodes(output.warnings)).toContain('EMPTY_DOCUMENT');
  });

  it('flags styles missing from the package', async () => {
    const output = await inspect(makeModule(), 'no-styles.docx');
    expect(output.result.metadata['styleCount']).toBe(0);
    expect(warningCodes(output.warnings)).toContain('STYLES_MISSING');
  });

  it('flags paragraphs referencing undefined styles', async () => {
    const output = await inspect(makeModule(), 'broken-style.docx');
    expect(warningCodes(output.warnings)).toContain('BROKEN_STYLE_REFERENCE');
  });

  it('flags internal relationships with missing targets', async () => {
    const output = await inspect(makeModule(), 'dangling.docx');
    expect(warningCodes(output.warnings)).toContain('DANGLING_RELATIONSHIP');
  });

  it('parses comments in stable numeric order', async () => {
    const output = await execute(makeModule(), 'comments.docx');
    expect(output.result.ir.comments.map((comment) => comment.id)).toEqual(['1', '2']);
    expect(output.result.ir.comments[0]?.author).toBe('Reviewer A');
    expect(warningCodes(output.warnings)).toContain('COMMENTS_PRESENT');
  });

  it('parses footnotes and endnotes while dropping separators', async () => {
    const output = await execute(makeModule(), 'footnotes.docx');
    expect(output.result.ir.counts).toMatchObject({ footnotes: 2, endnotes: 1 });
    // 分隔符（type != normal）不应进入 IR。
    expect(output.result.ir.footnotes).toHaveLength(3);
    expect(warningCodes(output.warnings)).toContain('FOOTNOTES_PRESENT');
  });
});

describe.skipIf(!hasPython)('edge: resource budgets', () => {
  it('fails with LIMIT_EXCEEDED when the block budget is enforced', async () => {
    const module = makeModule({ limits: { maxBlocks: 5 } });
    await expect(execute(module, 'many-blocks.docx')).rejects.toMatchObject({
      code: 'LIMIT_EXCEEDED',
    });
  });

  it('degrades gracefully when enforcement is off', async () => {
    const module = makeModule({
      limits: { maxBlocks: 5 },
      featureFlags: { enforceLimits: false },
    });
    const output = await execute(module, 'many-blocks.docx');
    expect(output.result.ir.counts.blocks).toBe(5);
    expect(warningCodes(output.warnings)).toContain('LIMIT_APPLIED');
  });
});

describe.skipIf(!hasPython)('edge: damaged and rejected artifacts', () => {
  it('rejects an OLE container as out of scope', async () => {
    await expect(inspect(makeModule(), 'legacy.doc')).rejects.toMatchObject({
      code: 'UNSUPPORTED_CONTAINER',
    });
  });

  it('rejects RTF as out of scope', async () => {
    await expect(inspect(makeModule(), 'sample.rtf')).rejects.toMatchObject({
      code: 'UNSUPPORTED_CONTAINER',
    });
  });

  it('rejects unrecognizable bytes as a format mismatch', async () => {
    await expect(inspect(makeModule(), 'plain.bin')).rejects.toMatchObject({
      code: 'FORMAT_MISMATCH',
    });
  });

  it('rejects a non-Word OOXML package', async () => {
    await expect(inspect(makeModule(), 'spreadsheet.docx')).rejects.toMatchObject({
      code: 'FORMAT_MISMATCH',
    });
  });

  it('rejects a package without content types', async () => {
    await expect(inspect(makeModule(), 'no-content-types.docx')).rejects.toMatchObject({
      code: 'FORMAT_MISMATCH',
    });
  });

  it('fails on a truncated archive', async () => {
    await expect(inspect(makeModule(), 'truncated.docx')).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('fails on a missing artifact', async () => {
    const module = makeModule();
    await expect(
      module.handlers.inspect({
        artifactRef: { id: 'missing', uri: 'fixtures/_generated/does-not-exist.docx' },
        operation: 'inspect',
        requestId: 'req-missing',
      }),
    ).rejects.toMatchObject({ code: 'ARTIFACT_NOT_FOUND' });
  });

  it('detects encryption and lets the policy deny it', async () => {
    const module = makeModule();
    const output = await inspect(module, 'encrypted.docx');
    expect(output.result.encrypted).toBe(true);

    await expect(
      module.handlers.inspect({
        artifactRef: artifactRef('encrypted.docx'),
        operation: 'inspect',
        requestId: 'req-enc',
        policy: { id: 'strict', allowEncrypted: false },
      }),
    ).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
  });

  it('honours a declared extension that contradicts the content', async () => {
    const module = makeModule();
    await expect(
      inspect(module, 'clean.docx', { options: { declaredExtension: 'xlsx' } }),
    ).rejects.toMatchObject({ code: 'FORMAT_MISMATCH' });
  });
});

describe.skipIf(!hasPython)('edge: verification against real documents', () => {
  it('fails policy checks for a document lacking the required structure', async () => {
    const module = makeModule();
    const empty = await verify(module, 'empty.docx', { requireHeadings: true, minParagraphs: 1 });
    expect(empty.result.ok).toBe(false);
    expect(
      empty.result.checks.filter((entry) => entry.status === 'fail').map((entry) => entry.id),
    ).toEqual(expect.arrayContaining(['policy.headings', 'policy.minParagraphs']));
  });

  it('detects unresolved style references and dangling relationships', async () => {
    const module = makeModule();
    const broken = await verify(module, 'broken-style.docx', { requireResolvedStyles: true });
    expect(broken.result.ok).toBe(false);

    const dangling = await verify(module, 'dangling.docx', {
      requireNoDanglingRelationships: true,
    });
    expect(dangling.result.ok).toBe(false);
  });

  it('passes all declared checks for a clean document', async () => {
    const module = makeModule();
    const output = await verify(module, 'clean.docx', {
      requireHeadings: true,
      minParagraphs: 2,
      requireStyles: true,
      requireResolvedStyles: true,
      requireNoDanglingRelationships: true,
      allowEncrypted: false,
    });
    expect(output.result.ok).toBe(true);
    expect(output.result.summary.failed).toBe(0);
  });
});

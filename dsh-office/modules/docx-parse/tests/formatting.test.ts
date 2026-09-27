import { describe, expect, it } from 'vitest';
import { createDocxParseModule } from '../src/index';
import { parseParseResult } from '../src/domain/docx-parse';
import type { DocxDualIR } from '../src/index';
import type { RawBlock } from '../src/domain/docx-parse';
import { FakeEngine, artifactRef, para, parseResult } from './support';

/**
 * 表现层观察里「首行缩进」这一项。
 *
 * 为什么不手搓 IR：`blockId`、锚点、对应表都是 mapper 推导出来的，手搓会把被测的
 * 那一段绕过去。这里用替身引擎喂**原始观察**，跑真实 mapper。
 */
async function observe(blocks: RawBlock[], formatting = true): Promise<DocxDualIR['formatting']> {
  const module = createDocxParseModule({ engine: new FakeEngine(parseResult({ blocks })) });
  try {
    const output = await module.handlers.execute({
      artifactRef: artifactRef('file:///tmp/indent.docx'),
      operation: 'execute',
      requestId: 'indent-observation',
      ...(formatting ? { options: { featureFlags: { parseFormatting: true } } } : {}),
    });
    return output.result.ir.content.formatting;
  } finally {
    await module.dispose();
  }
}

const formattingOf = (indent: { firstLineTwips?: number; firstLineChars?: number } | undefined) => ({
  alignment: null,
  ...(indent ? { indent } : {}),
  runs: [{ text: '正文', size: 12, eastAsia: '宋体' }],
});

describe('表现层观察：首行缩进', () => {
  it('把解析过样式链的缩进带到 IR 上，0 也照样带', async () => {
    const blocks: RawBlock[] = [
      para('缩进两字的正文', 0, { formatting: formattingOf({ firstLineTwips: 480, firstLineChars: 200 }) }),
      para('齐左的信息行', 1, { formatting: formattingOf({ firstLineTwips: 0, firstLineChars: 0 }) }),
    ];
    const formatting = await observe(blocks);
    expect(formatting?.paragraphs.map((entry) => entry.indent)).toEqual([
      { firstLineTwips: 480, firstLineChars: 200 },
      { firstLineTwips: 0, firstLineChars: 0 },
    ]);
  });

  it('观察里没有缩进就不写，0 不能当成缺省值', async () => {
    const blocks: RawBlock[] = [para('没有缩进观察的段落', 0, { formatting: formattingOf(undefined) })];
    const formatting = await observe(blocks);
    expect(formatting?.paragraphs).toHaveLength(1);
    expect(formatting?.paragraphs[0]?.indent).toBeUndefined();
    expect('indent' in (formatting?.paragraphs[0] ?? {})).toBe(false);
  });

  it('开关关闭时整节观察缺席，缩进也就无从谈起', async () => {
    const blocks: RawBlock[] = [para('正文', 0, { formatting: formattingOf({ firstLineTwips: 480 }) })];
    expect(await observe(blocks, false)).toBeUndefined();
  });

  it('缩进不是数字时被信任边界拒绝，而不是悄悄当成 0', () => {
    // 载荷来自子进程，是模块的信任边界；畸形输出只能变成受控的协议错误。
    const payload = parseResult({
      blocks: [
        para('坏缩进', 0, {
          formatting: { alignment: null, indent: { firstLineTwips: '两字' }, runs: [{ text: '正文' }] } as never,
        }),
      ],
    });
    expect(() => parseParseResult(payload)).toThrow(/firstLineTwips/u);
  });
});

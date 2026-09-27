/**
 * 回归测试 —— 锁住「已经正确的行为」，防止后续改动悄悄破坏它们。
 *
 * 关注三类不变量：
 *   1. 确定性：同一份输入必须得到逐字节一致的 IR（否则 diff / 快照 / 缓存全部失效）；
 *   2. 接口稳定性：换掉底层引擎，公开输出形状不能变（spec 完成标准第四条）；
 *   3. 序列化：公开输出在任何情况下都可 JSON 往返。
 */
import { describe, expect, it } from 'vitest';

import { assertParseUsable, computeCounts, toDocxParseIR, toWarnings } from '../src/internal';
import { resolveConfig } from '../src/config';
import {
  artifactRef,
  expectJsonRoundTrip,
  fakeEngine,
  fakeParseResult,
  makeModule,
  makeModuleWithEngine,
  pythonAvailable,
} from './support';

const hasPython = pythonAvailable();

describe.skipIf(!hasPython)('regression: deterministic output', () => {
  it('produces byte-identical IR across repeated runs', async () => {
    const module = makeModule();
    const run = () =>
      module.handlers.execute({
        artifactRef: artifactRef('clean.docx'),
        operation: 'execute',
        requestId: 'req-determinism',
      });

    const first = await run();
    const second = await run();
    // 只比较 result：telemetry 里的耗时天然会变化。
    expect(JSON.stringify(second.result)).toBe(JSON.stringify(first.result));
  });

  it('orders warnings deterministically for the same document', async () => {
    const module = makeModule();
    const run = () =>
      module.handlers.inspect({
        artifactRef: artifactRef('broken-style.docx'),
        operation: 'inspect',
        requestId: 'req-warnings',
      });

    const first = await run();
    const second = await run();
    expect(second.warnings).toEqual(first.warnings);
  });

  it('keeps the IR envelope keys stable', async () => {
    const output = await makeModule().handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-keys',
    });
    expect(Object.keys(output).sort()).toEqual([
      'artifacts',
      'moduleId',
      'operation',
      'requestId',
      'result',
      'telemetry',
      'warnings',
    ]);
    expect(Object.keys(output.result.ir).sort()).toEqual([
      'blocks',
      'comments',
      'counts',
      'footnotes',
      'metadata',
      'outline',
      'profile',
      'relationships',
      'styles',
    ]);
  });
});

describe('regression: engine replaceability', () => {
  it('keeps the public output shape identical when the engine is swapped', async () => {
    const result = fakeParseResult({
      blocks: [
        {
          kind: 'paragraph',
          paragraph: {
            index: 0,
            text: 'Swapped',
            styleId: 'Heading1',
            styleName: 'heading 1',
            outlineLevel: 0,
            headingLevel: 1,
            alignment: null,
            listItem: false,
            runs: [],
          },
        },
      ],
    });

    const fake = makeModuleWithEngine(fakeEngine(result));
    const output = await fake.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-swap',
    });

    // 引擎换了，但 IR 的形状与派生结果必须完全由 mapper 决定。
    const config = resolveConfig();
    assertParseUsable(result, config);
    const expected = toDocxParseIR(result);
    expect(output.result.ir).toEqual(expected);
    expect(toWarnings(result, {}, config).map((warning) => warning.code)).toEqual(
      output.warnings.map((warning) => warning.code),
    );
    expect(computeCounts(result).blocks).toBe(1);
    expectJsonRoundTrip(output);
  });
});

describe('regression: serialization guarantees', () => {
  it('serializes every public artifact of a module call', async () => {
    const module = makeModuleWithEngine(fakeEngine(fakeParseResult()));
    const inspectOutput = await module.handlers.inspect({
      artifactRef: artifactRef('clean.docx'),
      operation: 'inspect',
      requestId: 'req-serial-inspect',
    });
    const executeOutput = await module.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-serial-execute',
    });
    const verifyOutput = await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-serial-verify',
      policy: { id: 'p' },
    });
    expectJsonRoundTrip(inspectOutput);
    expectJsonRoundTrip(executeOutput);
    expectJsonRoundTrip(verifyOutput);
  });
});

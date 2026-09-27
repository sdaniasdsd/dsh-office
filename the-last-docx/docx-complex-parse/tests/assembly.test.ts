/**
 * 装配层的规格：模块「注册进去」之后，Profile 看到的东西必须是对的。
 *
 * 这里刻意注入一个假引擎：装配层的职责是拼接与翻译，不该依赖真的起 Python 进程。
 * 引擎的行为差异由 `domain.test.ts` 与黑盒测试覆盖。
 */
import { describe, expect, it } from 'vitest';

import {
  createDocxComplexParseModule,
  DOCX_COMPLEX_PARSE_DEFINITION,
  register,
  type ProfileRegistryLike,
} from '../src/index';
import { DocxComplexParseError } from '../src/errors';
import type { DocxEngine, ParseRequest } from '../src/engine/adapter';
import type { RawParseResult } from '../src/domain/docx-complex-parse';
import type { ArtifactRef } from 'office-core';
import { cell, paragraph, rawResult, row, table } from './support';

/** 一个只回答固定结果的假引擎，并记录调用参数供断言。 */
function fakeEngine(
  result: RawParseResult | Error,
  calls: ParseRequest[] = [],
): DocxEngine {
  return {
    name: 'fake',
    async parse(request: ParseRequest): Promise<RawParseResult> {
      calls.push(request);
      if (result instanceof Error) throw result;
      return result;
    },
    async dispose(): Promise<void> {
      // 无状态。
    },
  };
}

const ARTIFACT: ArtifactRef = {
  id: 'artifact-1',
  uri: 'C:/docs/report.docx',
  label: 'report.docx',
};

/** 调用一个 handler 并拿到结果，同时把抛出的错误收敛成值便于断言。 */
async function run<T>(fn: () => Promise<T>): Promise<{ value?: T; error?: unknown }> {
  try {
    return { value: await fn() };
  } catch (error) {
    return { error };
  }
}

describe('模块定义', () => {
  it('身份、能力与依赖与 module.json 一致', () => {
    expect(DOCX_COMPLEX_PARSE_DEFINITION.id).toBe('docx-complex-parse');
    expect(DOCX_COMPLEX_PARSE_DEFINITION.version).toBe('0.1.0');
    expect(DOCX_COMPLEX_PARSE_DEFINITION.profileGroup).toBe('DOCX');
    expect(DOCX_COMPLEX_PARSE_DEFINITION.capabilities).toEqual(['inspect', 'execute', 'verify']);
  });

  it('三个 handler 都存在且是函数', () => {
    const module = createDocxComplexParseModule({ engine: fakeEngine(rawResult([])) });
    expect(typeof module.handlers.inspect).toBe('function');
    expect(typeof module.handlers.execute).toBe('function');
    expect(typeof module.handlers.verify).toBe('function');
  });

  it('注册到 registry 时按契约调用 registerModule', async () => {
    const seen: unknown[] = [];
    const registry: ProfileRegistryLike = {
      registerModule: (module) => {
        seen.push(module);
      },
    };
    const module = await register(registry, { engine: fakeEngine(rawResult([])) });
    expect(seen).toHaveLength(1);
    expect(seen[0]).toBe(module);
  });
});

describe('inspect', () => {
  it('产出 FormatProfile，并回填扩展名与包级计数', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([paragraph('hi')], {
          parts: [
            { name: 'word/document.xml', sizeBytes: 2048, compressedSize: 900 },
            { name: 'word/styles.xml', sizeBytes: 700, compressedSize: 300 },
          ],
          relationships: [
            {
              sourcePart: 'word/document.xml',
              id: 'rId1',
              type: 't',
              target: 'styles.xml',
              targetMode: 'Internal',
            },
          ],
        }),
      ),
    });

    const out = await module.handlers.inspect({
      artifactRef: ARTIFACT,
      operation: 'inspect',
      requestId: 'req-1',
    });

    expect(out.moduleId).toBe('docx-complex-parse');
    expect(out.operation).toBe('inspect');
    expect(out.result.format).toBe('docx');
    // 扩展名来自 artifactRef.label，而不是内容——两者互补。
    expect(out.result.extension).toBe('docx');
    expect(out.result.confidence).toBe(1);
    expect(out.result.features['hasMacros']).toBe(false);
    expect(out.result.metadata['partCount']).toBe(2);
    expect(out.result.metadata['relationshipCount']).toBe(1);
    // inspect 不产出文件。
    expect(out.artifacts).toEqual([]);
    expect(out.telemetry?.relationshipCount).toBe(1);
  });

  it('带宏时给出 UNSUPPORTED_CONTENT 告警，但不拒绝', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([paragraph('hi')], {
          profile: { format: 'docm' },
          indicators: {
            macros: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 4096 }],
          },
        }),
      ),
    });

    const out = await module.handlers.inspect({
      artifactRef: ARTIFACT,
      operation: 'inspect',
      requestId: 'req-2',
      options: { declaredExtension: 'docm' },
    });

    expect(out.result.features['hasMacros']).toBe(true);
    expect(out.warnings.map((warning) => warning.code)).toContain('UNSUPPORTED_CONTENT');
  });

  it('策略拒绝时以 SAFETY_POLICY_DENIED 快速失败', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([paragraph('hi')], {
          profile: { format: 'docm' },
          indicators: {
            macros: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 128 }],
          },
        }),
      ),
    });

    const { error } = await run(() =>
      module.handlers.inspect({
        artifactRef: ARTIFACT,
        operation: 'inspect',
        requestId: 'req-3',
        options: { declaredExtension: 'docm' },
        policy: { id: 'strict', allowMacros: false },
      }),
    );

    expect(error).toBeInstanceOf(DocxComplexParseError);
    expect((error as DocxComplexParseError).code).toBe('SAFETY_POLICY_DENIED');
  });

  it('声明 docx 而内容是 docm 时判为 FORMAT_MISMATCH', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(rawResult([paragraph('hi')], { profile: { format: 'docm' } })),
    });

    const { error } = await run(() =>
      module.handlers.inspect({
        artifactRef: ARTIFACT,
        operation: 'inspect',
        requestId: 'req-4',
        options: { declaredExtension: 'docx' },
      }),
    );

    expect((error as DocxComplexParseError).code).toBe('FORMAT_MISMATCH');
  });

  it('verifyAfterInspect 附上结构自检报告，且不会因策略拒绝', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(rawResult([paragraph('hi')])),
    });

    const out = await module.handlers.inspect({
      artifactRef: ARTIFACT,
      operation: 'inspect',
      requestId: 'req-5',
      options: { verifyAfterInspect: true },
    });

    expect(out.verification?.ok).toBe(true);
    // 没声明任何要求的策略下，策略类检查只能是 skip。
    expect(out.verification?.summary.skipped).toBeGreaterThan(0);
    expect(out.verification?.policyId).toBe('inspect.self-check');
  });
});

describe('execute', () => {
  it('产出 IR：包级事实 + 展开后的逻辑网格', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([
          paragraph('hello'),
          table([row([cell('a', 2), cell('b')])], 3),
        ]),
      ),
    });

    const out = await module.handlers.execute({
      artifactRef: ARTIFACT,
      operation: 'execute',
      requestId: 'req-6',
    });

    const { ir, artifact } = out.result;
    expect(ir.profile.format).toBe('docx');
    expect(ir.parts).toHaveLength(1);
    // 展开后列数按 tblGrid（3）而不是按格子数（2）。
    expect(ir.content.blocks[1]?.table?.columnCount).toBe(3);
    expect(ir.content.blocks[1]?.table?.grid[0]).toHaveLength(3);
    // artifact 未声明媒体类型时由内容补上。
    expect(artifact.mediaType).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    );
    expect(out.telemetry?.tableCount).toBe(1);
    expect(out.telemetry?.pageGeometryAvailable).toBe(false);
  });

  it('网格自相矛盾时把 MERGE_INCONSISTENT 报成告警', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([table([row([cell('orphan', 1, 'continue')])], 1)]),
      ),
    });

    const out = await module.handlers.execute({
      artifactRef: ARTIFACT,
      operation: 'execute',
      requestId: 'req-7',
    });

    const codes = out.warnings.map((warning) => warning.code);
    expect(codes).toContain('MERGE_INCONSISTENT');
    // 低置信度结论同时广播，但只是 info，不阻断。
    expect(codes).toContain('LOW_CONFIDENCE');
    expect(out.warnings.every((warning) => warning.code !== 'VERIFICATION_FAILED')).toBe(true);
  });

  it('引擎抛错时收敛为本模块错误', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(new Error('boom')),
    });

    const { error } = await run(() =>
      module.handlers.execute({ artifactRef: ARTIFACT, operation: 'execute', requestId: 'req-8' }),
    );

    expect(error).toBeInstanceOf(DocxComplexParseError);
    expect((error as DocxComplexParseError).code).toBe('ENGINE_FAILED');
  });

  it('引擎抛出的模块错误原样透出，不被重新包装', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(new DocxComplexParseError('ARTIFACT_NOT_FOUND', 'no such file')),
    });

    const { error } = await run(() =>
      module.handlers.execute({ artifactRef: ARTIFACT, operation: 'execute', requestId: 'req-9' }),
    );

    expect((error as DocxComplexParseError).code).toBe('ARTIFACT_NOT_FOUND');
  });

  it('调用级覆盖生效，且不污染模块级配置', async () => {
    const calls: ParseRequest[] = [];
    const module = createDocxComplexParseModule({
      config: { limits: { maxBlocks: 100 } },
      engine: fakeEngine(rawResult([paragraph('hi')]), calls),
    });

    await module.handlers.execute({
      artifactRef: ARTIFACT,
      operation: 'execute',
      requestId: 'req-10',
      options: { limits: { maxBlocks: 5 } },
    });
    expect(calls[0]?.limits.maxBlocks).toBe(5);

    await module.handlers.execute({
      artifactRef: ARTIFACT,
      operation: 'execute',
      requestId: 'req-11',
    });
    expect(calls[1]?.limits.maxBlocks).toBe(100);
  });

  it('超出预算且限额开启时失败；关闭时降级为 LIMIT_APPLIED 告警', async () => {
    const raw = rawResult([paragraph('a'), paragraph('b', 2)]);
    const strict = createDocxComplexParseModule({
      config: { limits: { maxBlocks: 1 } },
      engine: fakeEngine(raw),
    });
    const strictResult = await run(() =>
      strict.handlers.execute({ artifactRef: ARTIFACT, operation: 'execute', requestId: 'req-12' }),
    );
    expect((strictResult.error as DocxComplexParseError).code).toBe('LIMIT_EXCEEDED');

    const lenient = createDocxComplexParseModule({
      config: { limits: { maxBlocks: 1 }, featureFlags: { enforceLimits: false } },
      engine: fakeEngine(raw),
    });
    const lenientResult = await lenient.handlers.execute({
      artifactRef: ARTIFACT,
      operation: 'execute',
      requestId: 'req-13',
    });
    expect(lenientResult.warnings.map((warning) => warning.code)).toContain('LIMIT_APPLIED');
  });
});

describe('verify', () => {
  it('策略通过时报告 ok，且附上 policyId', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(rawResult([paragraph('hi')])),
    });

    const out = await module.handlers.verify({
      artifactRef: ARTIFACT,
      operation: 'verify',
      requestId: 'req-14',
      policy: { id: 'p1', allowMacros: false, allowEncrypted: false },
    });

    expect(out.result.ok).toBe(true);
    expect(out.result.policyId).toBe('p1');
    expect(out.verification).toBe(out.result);
  });

  it('策略要求未满足时报告 not ok，并广播 VERIFICATION_FAILED', async () => {
    const module = createDocxComplexParseModule({
      engine: fakeEngine(
        rawResult([paragraph('hi')], {
          profile: { format: 'docm' },
          indicators: {
            macros: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 64 }],
          },
        }),
      ),
    });

    const out = await module.handlers.verify({
      artifactRef: ARTIFACT,
      operation: 'verify',
      requestId: 'req-15',
      options: { declaredExtension: 'docm' },
      policy: { id: 'p2', allowMacros: false },
    });

    expect(out.result.ok).toBe(false);
    const failed = out.result.checks.filter((entry) => entry.status === 'fail');
    expect(failed.map((entry) => entry.id)).toContain('policy.macros');
    expect(out.warnings.map((warning) => warning.code)).toContain('VERIFICATION_FAILED');
  });
});

describe('dispose', () => {
  it('释放引擎且可重复调用', async () => {
    let disposed = 0;
    const engine = fakeEngine(rawResult([]));
    const module = createDocxComplexParseModule({
      engine: { ...engine, dispose: async () => { disposed += 1; } },
    });

    await module.dispose();
    await module.dispose();
    expect(disposed).toBe(2);
  });
});

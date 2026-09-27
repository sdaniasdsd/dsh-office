/**
 * 边界与拒绝场景测试。
 *
 * spec 完成标准要求「至少覆盖正常、边界、损坏或拒绝场景」，本文件就是这三类
 * 场景的集中体现。每个用例都显式声明「输入 → 期望的错误码或告警码」，
 * 从而把模块的失败语义固化成可执行的规范。
 */
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

import { beforeAll, describe, expect, it } from 'vitest';

import {
  DocxInspectError,
  createDocxInspectModule,
  isDocxInspectError,
} from '../src/index';
import type { DocxEngine, DocxInspectModule, JsonObject, ProbeResult } from '../src/index';

const FIXTURES_DIR = fileURLToPath(new URL('../fixtures/_generated/', import.meta.url));

/** 拼出 fixture 的绝对路径。 */
function fixture(name: string): string {
  return join(FIXTURES_DIR, name);
}

/** 让 requestId 可预测，便于断言输出。 */
let sequence = 0;
function nextRequestId(prefix: string): string {
  sequence += 1;
  return `${prefix}-${sequence}`;
}

let module: DocxInspectModule;

beforeAll(() => {
  // 每个测试文件用独立实例，避免相互影响。
  module = createDocxInspectModule();
});

/** 断言某次调用以指定的错误码失败，并返回该错误以便进一步检查。 */
async function expectFailureCode(
  run: () => Promise<unknown>,
  code: string,
): Promise<DocxInspectError> {
  try {
    await run();
    throw new Error(`expected failure ${code}, but the call succeeded`);
  } catch (error) {
    expect(isDocxInspectError(error)).toBe(true);
    const typed = error as DocxInspectError;
    expect(typed.code).toBe(code);
    return typed;
  }
}

describe('edge cases: accepted inputs', () => {
  it('inspects a well-formed document with full confidence', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('clean'),
      operation: 'inspect',
      artifactRef: { id: 'a1', uri: fixture('clean.docx') },
    });
    expect(output.result.format).toBe('docx');
    expect(output.result.container).toBe('zip');
    expect(output.result.encrypted).toBe(false);
    expect(output.result.confidence).toBe(1);
    expect(output.warnings).toEqual([]);
    // 探测不产生新文件。
    expect(output.artifacts).toEqual([]);
    expect(output.moduleId).toBe('docx-inspect');
  });

  it('flags an empty but structurally valid document without failing', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('empty'),
      operation: 'inspect',
      artifactRef: { id: 'a2', uri: fixture('empty.docx') },
    });
    // 空文档是「告警」而非「错误」：它能被正常读取，只是没有内容。
    const codes = output.warnings.map((warning) => warning.code);
    expect(codes).toContain('EMPTY_DOCUMENT');
    expect(output.result.features['hasMacros']).toBe(false);
  });
});

describe('edge cases: security indicators', () => {
  it('reports macros and the macro-enabled variant', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('macro'),
      operation: 'inspect',
      artifactRef: { id: 'a3', uri: fixture('macro.docm') },
    });
    expect(output.result.format).toBe('docm');
    expect(output.result.features['hasMacros']).toBe(true);
    const macroWarning = output.warnings.find((warning) => warning.code === 'MACROS_PRESENT');
    expect(macroWarning).toBeDefined();
    expect(macroWarning?.details?.['count']).toBe(1);
  });

  it('reports external targets with categories and DDE fields', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('external'),
      operation: 'inspect',
      artifactRef: { id: 'a4', uri: fixture('external.docx') },
    });
    const codes = output.warnings.map((warning) => warning.code);
    expect(codes).toContain('EXTERNAL_REFERENCES_PRESENT');
    expect(codes).toContain('DDE_FIELDS_PRESENT');
    expect(output.result.features['hasExternalLinks']).toBe(true);
    const external = output.warnings.find(
      (warning) => warning.code === 'EXTERNAL_REFERENCES_PRESENT',
    );
    // 三个外链分别是 远程模板 / 超链接 / 外链图片。
    expect(external?.details?.['categories']).toEqual([
      'attachedTemplate',
      'externalImage',
      'hyperlink',
    ]);
  });

  it('reports embedded objects, altChunks and ActiveX separately', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('embedded'),
      operation: 'inspect',
      artifactRef: { id: 'a5', uri: fixture('embedded.docx') },
    });
    const codes = output.warnings.map((warning) => warning.code);
    // 三者语义不同，必须各自出告警，不能混为一谈。
    expect(codes).toContain('EMBEDDED_OBJECTS_PRESENT');
    expect(codes).toContain('ALT_CHUNKS_PRESENT');
    expect(codes).toContain('ACTIVE_X_PRESENT');
    expect(output.result.features['hasEmbeddedObjects']).toBe(true);
    expect(output.result.features['hasAltChunks']).toBe(true);
    expect(output.result.features['hasActiveX']).toBe(true);
  });

  it('reports an encrypted package', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('encrypted'),
      operation: 'inspect',
      artifactRef: { id: 'a6', uri: fixture('encrypted.docx') },
    });
    expect(output.result.encrypted).toBe(true);
    expect(output.warnings.map((warning) => warning.code)).toContain('ENCRYPTED_ARTIFACT');
  });

  it('honours a disabled feature flag', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('nomacro'),
      operation: 'inspect',
      artifactRef: { id: 'a7', uri: fixture('macro.docm') },
      // 关闭宏探测后，结论应当反映「本次没有探测宏」，而不是「文档没有宏」。
      options: { featureFlags: { detectMacros: false } },
    });
    expect(output.warnings.map((warning) => warning.code)).not.toContain('MACROS_PRESENT');
    expect(output.result.features['hasMacros']).toBe(false);
  });
});

describe('edge cases: rejected inputs', () => {
  it('rejects a non-Word OOXML package with FORMAT_MISMATCH', async () => {
    const error = await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('xlsx'),
          operation: 'inspect',
          artifactRef: { id: 'b1', uri: fixture('not-word.docx') },
        }),
      'FORMAT_MISMATCH',
    );
    expect(error.details['documentKind']).toBe('spreadsheetml');
  });

  it('rejects a corrupted archive with PARSE_FAILED', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('truncated'),
          operation: 'inspect',
          artifactRef: { id: 'b2', uri: fixture('truncated.docx') },
        }),
      'PARSE_FAILED',
    );
  });

  it('rejects an OLE container with UNSUPPORTED_CONTAINER', async () => {
    const error = await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('ole'),
          operation: 'inspect',
          artifactRef: { id: 'b3', uri: fixture('legacy.doc') },
        }),
      'UNSUPPORTED_CONTAINER',
    );
    expect(error.details['container']).toBe('ole');
  });

  it('rejects an RTF stream with UNSUPPORTED_CONTAINER', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('rtf'),
          operation: 'inspect',
          artifactRef: { id: 'b4', uri: fixture('sample.rtf') },
        }),
      'UNSUPPORTED_CONTAINER',
    );
  });

  it('rejects unrecognized bytes with FORMAT_MISMATCH', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('plain'),
          operation: 'inspect',
          artifactRef: { id: 'b5', uri: fixture('plain.bin') },
        }),
      'FORMAT_MISMATCH',
    );
  });

  it('rejects a ZIP without [Content_Types].xml', async () => {
    // 缺 Content_Types 意味着 documentKind 无法判定，因此归为格式不匹配。
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('noct'),
          operation: 'inspect',
          artifactRef: { id: 'b6', uri: fixture('no-content-types.zip') },
        }),
      'FORMAT_MISMATCH',
    );
  });

  it('rejects a missing artifact with ARTIFACT_NOT_FOUND', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('missing'),
          operation: 'inspect',
          artifactRef: { id: 'b7', uri: fixture('does-not-exist.docx') },
        }),
      'ARTIFACT_NOT_FOUND',
    );
  });

  it('rejects a directory with ARTIFACT_NOT_FOUND', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('dir'),
          operation: 'inspect',
          artifactRef: { id: 'b8', uri: FIXTURES_DIR },
        }),
      'ARTIFACT_NOT_FOUND',
    );
  });
});

describe('edge cases: declared hint mismatches', () => {
  it('rejects an unrelated declared extension', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('hint-ext'),
          operation: 'inspect',
          artifactRef: { id: 'c1', uri: fixture('clean.docx') },
          options: { declaredExtension: 'pdf' },
        }),
      'FORMAT_MISMATCH',
    );
  });

  it('rejects an unrelated declared MIME type', async () => {
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('hint-mime'),
          operation: 'inspect',
          artifactRef: { id: 'c2', uri: fixture('clean.docx') },
          options: { declaredMimeType: 'application/pdf' },
        }),
      'FORMAT_MISMATCH',
    );
  });

  it('treats a variant mismatch as a soft conflict', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('hint-soft'),
      operation: 'inspect',
      artifactRef: { id: 'c3', uri: fixture('clean.docx') },
      // 都是 Word 扩展名，只是变体不同：告警而非拒绝。
      options: { declaredExtension: 'docm' },
    });
    expect(output.warnings.map((warning) => warning.code)).toContain('DECLARED_HINT_MISMATCH');
    expect(output.result.features['declaredHintsCompatible']).toBe(true);
  });

  it('ignores a generic content type', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('hint-generic'),
      operation: 'inspect',
      artifactRef: { id: 'c4', uri: fixture('clean.docx') },
      // 浏览器/网关常这样发送，属于「未表态」。
      options: { declaredMimeType: 'application/octet-stream' },
    });
    expect(output.warnings).toEqual([]);
  });
});

describe('edge cases: resource limits', () => {
  it('fails with LIMIT_EXCEEDED when the entry budget is exceeded', async () => {
    // clean.docx 有 3 个条目，把上限设为 2 即可稳定触发。
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('limit-entries'),
          operation: 'inspect',
          artifactRef: { id: 'd1', uri: fixture('clean.docx') },
          options: { limits: { maxArchiveEntries: 2 } },
        }),
      'LIMIT_EXCEEDED',
    );
  });

  it('fails with LIMIT_EXCEEDED on a decompression bomb', async () => {
    // bomb.docx 内含一个 4 MiB 的高度可压缩条目；把单条目上限压到 64 KiB。
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('limit-bomb'),
          operation: 'inspect',
          artifactRef: { id: 'd2', uri: fixture('bomb.docx') },
          options: { limits: { maxEntryUncompressedBytes: 64 * 1024 } },
        }),
      'LIMIT_EXCEEDED',
    );
  });

  it('downgrades the same overflow to a warning when enforcement is off', async () => {
    const output = await module.handlers.inspect({
      requestId: nextRequestId('limit-soft'),
      operation: 'inspect',
      artifactRef: { id: 'd3', uri: fixture('clean.docx') },
      options: {
        limits: { maxArchiveEntries: 2 },
        featureFlags: { enforceLimits: false },
      },
    });
    // 关闭强制后仍产出结果，但必须留下「预算被应用过」的痕迹。
    expect(output.result.format).toBe('docx');
    expect(output.warnings.map((warning) => warning.code)).toContain('LIMIT_APPLIED');
  });
});

describe('edge cases: engine failures', () => {
  it('reports ENGINE_TIMEOUT when the probe exceeds its budget', async () => {
    // 1 毫秒必然不够启动解释器与读取文件。
    await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('timeout'),
          operation: 'inspect',
          artifactRef: { id: 'e1', uri: fixture('clean.docx') },
          options: { timeoutMs: 1 },
        }),
      'ENGINE_TIMEOUT',
    );
  });

  it('reports ENGINE_UNAVAILABLE when the interpreter is missing', async () => {
    const broken = createDocxInspectModule({
      config: { engine: { pythonPath: 'definitely-not-a-python-binary' } },
    });
    try {
      await expectFailureCode(
        () =>
          broken.handlers.inspect({
            requestId: nextRequestId('nointerp'),
            operation: 'inspect',
            artifactRef: { id: 'e2', uri: fixture('clean.docx') },
          }),
        'ENGINE_UNAVAILABLE',
      );
    } finally {
      await broken.dispose();
    }
  });
});

describe('edge cases: policy denial during inspect', () => {
  it('fails with SAFETY_POLICY_DENIED when the supplied policy rejects macros', async () => {
    const error = await expectFailureCode(
      () =>
        module.handlers.inspect({
          requestId: nextRequestId('deny'),
          operation: 'inspect',
          artifactRef: { id: 'f1', uri: fixture('macro.docm') },
          policy: { id: 'no-macros', allowMacros: false },
        }),
      'SAFETY_POLICY_DENIED',
    );
    const details = error.details as JsonObject;
    expect(details['policyId']).toBe('no-macros');
    expect(details['failedChecks']).toContain('policy.macros');
  });
});

/* -------------------------------------------------------------------------- */
/* 以下用例通过注入「假引擎」来隔离被测逻辑：                                    */
/* 宏语义分析属于可选增强（依赖 oletools），若直接跑真实引擎，                      */
/* 测试结果就会随环境是否安装 oletools 而漂移。注入预设的 ProbeResult              */
/* 既能稳定验证映射逻辑，也顺带证明了「引擎可替换而不改公开接口」。                  */
/* -------------------------------------------------------------------------- */

/** 构造一个字段齐全的 `ProbeResult`，只用于注入假引擎。 */
function fakeProbeResult(overrides: Partial<ProbeResult> = {}): ProbeResult {
  return {
    probeVersion: 1,
    xmlBackend: 'lxml',
    container: 'zip',
    extension: 'docx',
    sizeBytes: 4096,
    sha256: 'a'.repeat(64),
    encrypted: false,
    documentKind: 'wordprocessingml',
    documentVariant: null,
    mediaType: null,
    partCount: 1,
    parts: [{ name: '[Content_Types].xml', sizeBytes: 256, compressedSize: 128 }],
    relationships: [],
    macroIndicators: [],
    macroAnalyses: [],
    externalReferences: [],
    embeddedObjects: [],
    ddeFields: [],
    limitHit: null,
    issues: [],
    error: null,
    ...overrides,
  };
}

/** 造一个只会返回预设结果的引擎，用于把上层逻辑与 Python 引擎解耦。 */
function engineReturning(result: ProbeResult): DocxEngine {
  return {
    name: 'fake-engine',
    async probe(): Promise<ProbeResult> {
      return result;
    },
    async dispose(): Promise<void> {
      // 无状态，无需释放。
    },
  };
}

describe('edge cases: macro code semantics (optional enhancement)', () => {
  it('maps a suspicious macro analysis into MACRO_SUSPICIOUS_CODE', async () => {
    // 模拟「自动执行 + 执行外部命令」的典型恶意宏结论。
    const injected = createDocxInspectModule({
      engine: engineReturning(
        fakeProbeResult({
          macroIndicators: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 2048 }],
          macroAnalyses: [
            {
              module: 'Module1',
              autoExec: true,
              write: false,
              execute: true,
              suspicious: true,
              flags: 'A-X',
              matches: ['AutoOpen', 'Shell'],
            },
          ],
        }),
      ),
    });

    const output = await injected.handlers.inspect({
      requestId: nextRequestId('macro-semantics'),
      operation: 'inspect',
      artifactRef: { id: 's1', uri: fixture('clean.docx') },
    });

    const codes = output.warnings.map((warning) => warning.code);
    // 「有宏」与「宏可疑」是两条独立信息，应同时存在。
    expect(codes).toContain('MACROS_PRESENT');
    expect(codes).toContain('MACRO_SUSPICIOUS_CODE');

    const suspicious = output.warnings.find(
      (warning) => warning.code === 'MACRO_SUSPICIOUS_CODE',
    );
    const details = suspicious?.details as JsonObject;
    // 证据必须随结论一起回传，否则上层无法审计。
    expect(details['modules']).toEqual(['Module1']);
    expect(details['flags']).toEqual(['A-X']);
    expect(details['matches']).toEqual(['AutoOpen', 'Shell']);
  });

  it('does not flag macros that only auto-execute without payload behaviour', async () => {
    // 只有 A 而没有 W/X：这是大量合法业务宏的形态，不应被判为可疑。
    // 该用例锁定的正是「不做见宏即报警」这一设计取舍。
    const injected = createDocxInspectModule({
      engine: engineReturning(
        fakeProbeResult({
          macroIndicators: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 1024 }],
          macroAnalyses: [
            {
              module: 'ThisDocument',
              autoExec: true,
              write: false,
              execute: false,
              suspicious: false,
              flags: 'A--',
              matches: ['Document_Open'],
            },
          ],
        }),
      ),
    });

    const output = await injected.handlers.inspect({
      requestId: nextRequestId('macro-benign'),
      operation: 'inspect',
      artifactRef: { id: 's2', uri: fixture('clean.docx') },
    });

    const codes = output.warnings.map((warning) => warning.code);
    expect(codes).toContain('MACROS_PRESENT');
    expect(codes).not.toContain('MACRO_SUSPICIOUS_CODE');
  });

  it('sorts evidence so repeated probes stay byte-identical', async () => {
    // 证据来自第三方库，顺序不可控；这里确认我们已在 mapper 里做了归一化。
    const injected = createDocxInspectModule({
      engine: engineReturning(
        fakeProbeResult({
          macroIndicators: [{ kind: 'vba', part: 'word/vbaProject.bin', sizeBytes: 512 }],
          macroAnalyses: [
            {
              module: 'Zeta',
              autoExec: true,
              write: true,
              execute: false,
              suspicious: true,
              flags: 'AW-',
              matches: ['Shell', 'AutoOpen'],
            },
            {
              module: 'Alpha',
              autoExec: true,
              write: false,
              execute: true,
              suspicious: true,
              flags: 'A-X',
              matches: ['CreateObject', 'Document_Open'],
            },
          ],
        }),
      ),
    });

    const output = await injected.handlers.inspect({
      requestId: nextRequestId('macro-sort'),
      operation: 'inspect',
      artifactRef: { id: 's3', uri: fixture('clean.docx') },
    });

    const suspicious = output.warnings.find(
      (warning) => warning.code === 'MACRO_SUSPICIOUS_CODE',
    );
    const details = suspicious?.details as JsonObject;
    expect(details['modules']).toEqual(['Alpha', 'Zeta']);
    expect(details['matches']).toEqual(['AutoOpen', 'CreateObject', 'Document_Open', 'Shell']);
  });

  it('keeps probing successful when the optional capability is enabled but absent', async () => {
    // 核心不变式：可选增强无论可用与否，都【不得】让主流程失败。
    // 该断言在「装了 oletools」与「没装 oletools」两种环境下都成立。
    const analysed = createDocxInspectModule({
      config: { featureFlags: { analyzeMacroCode: true } },
    });

    const output = await analysed.handlers.inspect({
      requestId: nextRequestId('macro-optional'),
      operation: 'inspect',
      artifactRef: { id: 's4', uri: fixture('macro.docm') },
    });

    // 存在性指标照常给出；语义分析则视依赖是否可用而定，两者互不阻塞。
    expect(output.result.format).toBe('docm');
    expect(output.warnings.map((warning) => warning.code)).toContain('MACROS_PRESENT');
  });
});

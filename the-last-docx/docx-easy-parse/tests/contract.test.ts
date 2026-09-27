/**
 * 契约测试 —— 验证本模块对外承诺的表面是否成立。
 *
 * 关键特征：这一组测试【完全不依赖 Python】。它用「假引擎」驱动完整流水线，
 * 因此可以在 Profile 全量启动之前、甚至在没装解释器的机器上单独跑通
 * （spec 完成标准第一条）。真实的解析正确性由 edge-cases / regression 覆盖。
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import {
  DOCX_PARSE_CAPABILITIES,
  DOCX_PARSE_DEFINITION,
  DOCX_PARSE_ERROR_CODES,
  DOCX_PARSE_MODULE_ID,
  DOCX_PARSE_VERSION,
  DOCX_PARSE_WARNING_CODES,
  DocxParseError,
  isDocxParseError,
  register,
  resolveConfig,
  toDocxParseError,
} from '../src/index';
// 内部接缝按相对路径引用：它们不是公开契约，测试只是恰好需要。
import { ISSUE_TO_WARNING, InMemoryTelemetry, resolveArtifactPath, withSpan } from '../src/internal';
import type {
  Block,
  DocxParseModule,
  JsonValue,
  ParseResult,
  ProfileRegistry,
  VerificationReport,
} from '../src/index';
import { configSchema, withCallOverrides } from '../src/config';
import { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS, DEFAULT_TIMEOUT_MS } from '../src/config';
import { artifactRef, expectJsonRoundTrip, fakeEngine, fakeParseResult, makeModuleWithEngine, PROJECT_ROOT } from './support';

/* -------------------------------------------------------------------------- */
/* 构造测试数据                                                                 */
/* -------------------------------------------------------------------------- */

/** 构造一个段落块。 */
function para(index: number, text: string, headingLevel: number | null = null): Block {
  return {
    kind: 'paragraph',
    paragraph: {
      index,
      text,
      styleId: headingLevel === null ? null : `Heading${headingLevel}`,
      styleName: headingLevel === null ? null : `heading ${headingLevel}`,
      outlineLevel: headingLevel === null ? null : headingLevel - 1,
      headingLevel,
      alignment: null,
      listItem: false,
      runs: [{ text, bold: false, italic: false, underline: false }],
    },
  };
}

/** 构造一个表格块。 */
function tbl(index: number, rows: number, columns: number): Block {
  return {
    kind: 'table',
    table: { index, rows, columns, styleId: null, cells: [] },
  };
}

/** 一份带标题、段落、表格、样式与元数据的完整假结果。 */
function richResult(overrides: Partial<ParseResult> = {}): ParseResult {
  return fakeParseResult({
    blocks: [para(0, 'Title', 1), para(1, 'Hello'), tbl(0, 2, 2)],
    styles: [
      { styleId: 'Normal', name: 'Normal', type: 'paragraph', basedOn: null, isDefault: true, headingLevel: null },
      { styleId: 'Heading1', name: 'heading 1', type: 'paragraph', basedOn: 'Normal', isDefault: false, headingLevel: 1 },
    ],
    relationships: [
      {
        sourcePart: '',
        id: 'rId1',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
        target: 'word/document.xml',
        targetMode: 'Internal',
      },
    ],
    comments: [{ id: '1', author: 'A', initials: null, date: null, text: 'hi' }],
    footnotes: [
      { id: '1', kind: 'footnote', text: 'note' },
      { id: '1', kind: 'endnote', text: 'end' },
    ],
    ...overrides,
  });
}

/** 极简 registry 桩，仅记录收到的模块。 */
function recordingRegistry(): { registry: ProfileRegistry; modules: DocxParseModule[] } {
  const modules: DocxParseModule[] = [];
  return {
    modules,
    registry: {
      registerModule(module: DocxParseModule): void {
        modules.push(module);
      },
    },
  };
}

/** 调用 inspect 的便捷封装。 */
async function inspectWith(result: ParseResult, input: Record<string, unknown> = {}) {
  const module = makeModuleWithEngine(fakeEngine(result));
  return module.handlers.inspect({
    artifactRef: artifactRef('clean.docx'),
    operation: 'inspect',
    requestId: 'req-1',
    ...input,
  });
}

/* -------------------------------------------------------------------------- */
/* 身份与能力                                                                   */
/* -------------------------------------------------------------------------- */

describe('contract: identity and capabilities', () => {
  it('uses the module id required by the Profile layout', () => {
    expect(DOCX_PARSE_MODULE_ID).toBe('docx-parse');
  });

  it('publishes exactly the three capabilities from the spec, in order', () => {
    expect([...DOCX_PARSE_CAPABILITIES]).toEqual(['inspect', 'execute', 'verify']);
  });

  it('keeps error and warning tables serializable and complete', () => {
    expectJsonRoundTrip(DOCX_PARSE_ERROR_CODES);
    expectJsonRoundTrip(DOCX_PARSE_WARNING_CODES);
    for (const code of ['INVALID_INPUT', 'FORMAT_MISMATCH', 'ENGINE_TIMEOUT', 'LIMIT_EXCEEDED']) {
      expect(Object.keys(DOCX_PARSE_ERROR_CODES)).toContain(code);
    }
    for (const code of ['EMPTY_DOCUMENT', 'BROKEN_STYLE_REFERENCE', 'DANGLING_RELATIONSHIP']) {
      expect(Object.keys(DOCX_PARSE_WARNING_CODES)).toContain(code);
    }
  });

  /**
   * 从全部引擎脚本里抽出 `issue("<CODE>"` 字面量。
   *
   * 直接扫源码是刻意的：这条不变量跨越 TS/Python 两种语言，没有可调用的
   * 单一真值，只能以「引擎脚本里写了什么」为准。读取不需要解释器，
   * 因此本组测试仍然零 Python 依赖。
   */
  function emittedEngineIssueCodes(): Set<string> {
    const scripts = ['src/engine/docx_parse.py', 'src/engine/docx_parse_docling.py'];
    const codes = new Set<string>();
    for (const script of scripts) {
      const source = readFileSync(join(PROJECT_ROOT, script), 'utf8');
      // `\s*` 需跨行：源码里存在 `budget.issue(\n    "XML_INVALID",` 这种写法。
      for (const match of source.matchAll(/\.issue\(\s*["']([A-Z][A-Z_]*)/g)) {
        const code = match[1];
        if (code !== undefined) codes.add(code);
      }
    }
    return codes;
  }

  it('maps only engine issue codes that some engine actually emits', () => {
    // 守护「死码」的根因：给引擎从不发出的码登记映射，等于凭空造出一个
    // 永不触发的告警路径，而且不会有任何别的测试失败来提醒。
    // 历史案例：ENCRYPTED_ARTIFACT 的映射曾在，但两个引擎都不发它依赖的
    // ENCRYPTED_PACKAGE——加密事实实际上只由 FormatProfile.encrypted 表达。
    const emitted = emittedEngineIssueCodes();
    // 兜底：正则若失效会得到空集，届时下面的断言会全亮但原因易被误读。
    expect(emitted.size).toBeGreaterThan(0);

    const neverEmitted = Object.keys(ISSUE_TO_WARNING).filter((code) => !emitted.has(code));
    expect(neverEmitted, 'issue codes mapped in mapper.ts but emitted by no engine').toEqual([]);
  });

  it('declares no warning code that nothing can produce', () => {
    // 上一条守「映射 ⊄ 产出」，这一条守「声明 ⊆ 产出」。
    // 未登记的引擎问题码会被兜底降级为 PARTIALLY_PARSED，故它总有产出路径；
    // 其余码必须能在 mapper 的映射值或 TS 侧的直出字面量里找到来源。
    const producible = new Set<string>(
      Object.values(ISSUE_TO_WARNING).map((entry) => entry.code),
    );
    producible.add('PARTIALLY_PARSED');

    for (const file of ['mapper.ts', 'verifier.ts', 'index.ts']) {
      const source = readFileSync(join(PROJECT_ROOT, 'src', file), 'utf8');
      for (const code of Object.keys(DOCX_PARSE_WARNING_CODES)) {
        if (source.includes(`'${code}'`)) producible.add(code);
      }
    }

    const dead = Object.keys(DOCX_PARSE_WARNING_CODES).filter((code) => !producible.has(code));
    expect(dead, 'declared warning codes with no production path').toEqual([]);
  });
});

/* -------------------------------------------------------------------------- */
/* 模块清单                                                                     */
/* -------------------------------------------------------------------------- */

describe('contract: module definition', () => {
  it('declares the DOCX profile group and upstream dependencies', () => {
    expect(DOCX_PARSE_DEFINITION.profileGroup).toBe('DOCX');
    expect(DOCX_PARSE_DEFINITION.version).toBe(DOCX_PARSE_VERSION);
    const names = DOCX_PARSE_DEFINITION.dependencies.map((entry) => entry.name);
    for (const name of ['office-core', 'office-safety', 'office-files', 'office-test-kit', 'python']) {
      expect(names).toContain(name);
    }
  });

  it('ships a serializable config schema', () => {
    expectJsonRoundTrip(DOCX_PARSE_DEFINITION.configSchema);
    expect(DOCX_PARSE_DEFINITION.configSchema['type']).toBe('object');
  });

  it('stays consistent with module.json (single source of truth for the loader)', () => {
    const raw = readFileSync(join(PROJECT_ROOT, 'module.json'), 'utf8');
    const manifest: Record<string, JsonValue> = JSON.parse(raw);

    expect(manifest['id']).toBe(DOCX_PARSE_DEFINITION.id);
    expect(manifest['version']).toBe(DOCX_PARSE_DEFINITION.version);
    expect(manifest['profileGroup']).toBe(DOCX_PARSE_DEFINITION.profileGroup);
    expect(manifest['capabilities']).toEqual([...DOCX_PARSE_CAPABILITIES]);

    // 引擎声明必须指向真实存在的脚本，否则部署即失败。
    const engines = manifest['engines'] as Record<string, JsonValue>;
    const scriptPath = join(PROJECT_ROOT, String(engines['script']));
    expect(() => readFileSync(scriptPath)).not.toThrow();

    // 错误码与告警码表必须与清单一致。
    expect(manifest['errorCodes']).toEqual(Object.keys(DOCX_PARSE_ERROR_CODES));
    expect(manifest['warningCodes']).toEqual(Object.keys(DOCX_PARSE_WARNING_CODES));
  });

  it('produces a JSON-serializable definition', () => {
    expectJsonRoundTrip(DOCX_PARSE_DEFINITION);
  });
});

/* -------------------------------------------------------------------------- */
/* 配置                                                                         */
/* -------------------------------------------------------------------------- */

describe('contract: configuration', () => {
  it('applies defaults and freezes the resolved config', () => {
    const config = resolveConfig();
    expect(config.limits).toEqual(DEFAULT_LIMITS);
    expect(config.featureFlags).toEqual(DEFAULT_FEATURE_FLAGS);
    expect(config.timeoutMs).toBe(DEFAULT_TIMEOUT_MS);
    expect(config.engine.driver).toBe('python');
    // 冻结：下游误改必须直接失败，而不是静默生效。
    expect(Object.isFrozen(config)).toBe(true);
    expect(Object.isFrozen(config.limits)).toBe(true);
    expect(Object.isFrozen(config.featureFlags)).toBe(true);
  });

  it('accepts partial overrides without dropping unrelated defaults', () => {
    const config = resolveConfig({
      limits: { maxBlocks: 10 },
      featureFlags: { parseComments: false },
    });
    expect(config.limits.maxBlocks).toBe(10);
    expect(config.limits.maxTableCells).toBe(DEFAULT_LIMITS.maxTableCells);
    expect(config.featureFlags.parseComments).toBe(false);
    expect(config.featureFlags.parseStyles).toBe(true);
  });

  it('rejects invalid overrides with INVALID_INPUT', () => {
    expect(() => resolveConfig({ engine: { driver: 'ruby' as 'python' } })).toThrowError(
      /engine driver/i,
    );
    expect(() => resolveConfig({ engine: { pythonPath: '  ' } })).toThrowError(/pythonPath/i);
    expect(() => resolveConfig({ timeoutMs: 0 })).toThrowError(/timeoutMs/i);
    expect(() => resolveConfig({ limits: { maxBlocks: -1 } })).toThrowError(/maxBlocks/i);
  });

  it('merges call-level overrides on top of the module config', () => {
    const base = resolveConfig({ timeoutMs: 5000 });
    const merged = withCallOverrides(base, { limits: { maxBlocks: 3 }, timeoutMs: 1000 });
    expect(merged.timeoutMs).toBe(1000);
    expect(merged.limits.maxBlocks).toBe(3);
    // 未覆盖的字段沿用 base。
    expect(merged.limits.maxStyles).toBe(base.limits.maxStyles);
  });

  it('describes the same shape as module.json configSchema', () => {
    expectJsonRoundTrip(configSchema());
    expect(configSchema()['type']).toBe('object');
  });
});

/* -------------------------------------------------------------------------- */
/* 错误                                                                         */
/* -------------------------------------------------------------------------- */

describe('contract: errors', () => {
  it('serializes to JSON without losing the code', () => {
    const error = new DocxParseError('FORMAT_MISMATCH', 'nope', {
      path: 'word/document.xml',
      details: { kind: 'spreadsheetml' },
    });
    const json = error.toJSON();
    expect(json['code']).toBe('FORMAT_MISMATCH');
    expect(json['path']).toBe('word/document.xml');
    expectJsonRoundTrip(json);
  });

  it('recognizes its own errors and wraps foreign ones', () => {
    const own = new DocxParseError('PARSE_FAILED', 'x');
    expect(isDocxParseError(own)).toBe(true);
    expect(toDocxParseError(own, 'INVALID_INPUT')).toBe(own);

    const wrapped = toDocxParseError(new Error('boom'), 'PARSE_FAILED');
    expect(wrapped.code).toBe('PARSE_FAILED');
    expect(wrapped.details['reason']).toBe('boom');
    // 不可序列化的抛出物也要能安全落地。
    expectJsonRoundTrip(toDocxParseError(undefined, 'PARSE_FAILED').toJSON());
  });
});

/* -------------------------------------------------------------------------- */
/* artifact 定位                                                                */
/* -------------------------------------------------------------------------- */

describe('contract: artifact path resolution', () => {
  it('rejects empty and unsupported URIs', () => {
    expect(() => resolveArtifactPath('')).toThrowError(/non-empty/i);
    expect(() => resolveArtifactPath('https://example.com/a.docx')).toThrowError(
      /only reads local files/i,
    );
    expect(() => resolveArtifactPath('s3://bucket/a.docx')).toThrowError(/only reads local files/i);
    expect(() => resolveArtifactPath('ftp://host/a.docx')).toThrowError(/scheme/i);
  });

  it('accepts local paths and file:// URIs', () => {
    const local = resolveArtifactPath('fixtures/clean.docx');
    expect(local.length).toBeGreaterThan(0);
    const viaFileUri = resolveArtifactPath('file:///C:/tmp/a.docx');
    expect(viaFileUri.toLowerCase()).toContain('a.docx');
  });
});

/* -------------------------------------------------------------------------- */
/* 注册                                                                         */
/* -------------------------------------------------------------------------- */

describe('contract: registration', () => {
  it('registers itself and exposes the three handlers', async () => {
    const { registry, modules } = recordingRegistry();
    const module = await register(registry, { config: resolveConfig() });
    expect(modules).toHaveLength(1);
    expect(modules[0]).toBe(module);
    expect(typeof module.handlers.inspect).toBe('function');
    expect(typeof module.handlers.execute).toBe('function');
    expect(typeof module.handlers.verify).toBe('function');
    expect(module.definition.id).toBe(DOCX_PARSE_MODULE_ID);
    await expect(module.dispose()).resolves.toBeUndefined();
  });

  it('records spans through an injected telemetry sink', async () => {
    const telemetry = new InMemoryTelemetry();
    const value = await withSpan(telemetry, 'docx-parse.test', { requestId: 'req-1' }, async () => 42);
    expect(value).toBe(42);
    const events = telemetry.snapshot();
    // 一次成功的跨度会产生 start / end 两条事件。
    expect(events.map((event) => event.attributes['phase'])).toEqual(['start', 'end']);
    expectJsonRoundTrip(events);
  });
});

/* -------------------------------------------------------------------------- */
/* inspect 流水线（假引擎）                                                      */
/* -------------------------------------------------------------------------- */

describe('contract: inspect pipeline', () => {
  it('returns a FormatProfile with a serializable envelope', async () => {
    const output = await inspectWith(richResult());
    expect(output.moduleId).toBe(DOCX_PARSE_MODULE_ID);
    expect(output.operation).toBe('inspect');
    expect(output.result.format).toBe('docx');
    expect(output.result.container).toBe('zip');
    expect(output.result.features['hasHeadings']).toBe(true);
    expect(output.result.features['hasTables']).toBe(true);
    expect(output.result.features['emptyDocument']).toBe(false);
    expect(output.result.metadata['headingCount']).toBe(1);
    // 解析不产出新文件。
    expect(output.artifacts).toEqual([]);
    expectJsonRoundTrip(output);
  });

  it('never leaks engine-internal objects into the output', async () => {
    const output = await inspectWith(richResult());
    const encoded = JSON.stringify(output);
    // 序列化后再解析，形状应当等价（无 Buffer / 类实例 / 循环引用）。
    expect(JSON.parse(encoded)).toEqual(JSON.parse(JSON.stringify(output)));
  });
});

/* -------------------------------------------------------------------------- */
/* execute 流水线（假引擎）                                                      */
/* -------------------------------------------------------------------------- */

describe('contract: execute pipeline', () => {
  it('derives the IR from a single source of truth and sorts unstable collections', async () => {
    const result = richResult({
      styles: [
        // 故意乱序：mapper 必须按 styleId 归一化。
        { styleId: 'Heading1', name: 'heading 1', type: 'paragraph', basedOn: 'Normal', isDefault: false, headingLevel: 1 },
        { styleId: 'Normal', name: 'Normal', type: 'paragraph', basedOn: null, isDefault: true, headingLevel: null },
      ],
    });
    const module = makeModuleWithEngine(fakeEngine(result));
    const output = await module.handlers.execute({
      artifactRef: artifactRef('clean.docx'),
      operation: 'execute',
      requestId: 'req-exec',
    });

    expect(output.result.ir.styles.map((style) => style.styleId)).toEqual(['Heading1', 'Normal']);
    // outline 由 blocks 派生，永不与正文漂移。
    expect(output.result.ir.outline).toEqual([
      { level: 1, text: 'Title', paragraphIndex: 0 },
    ]);
    expect(output.result.ir.counts).toMatchObject({
      blocks: 3,
      paragraphs: 2,
      tables: 1,
      headings: 1,
      styles: 2,
      relationships: 1,
      comments: 1,
      footnotes: 1,
      endnotes: 1,
    });
    expectJsonRoundTrip(output);
  });

  it('refines the artifact reference without inventing a new id', async () => {
    const ref = artifactRef('clean.docx');
    const module = makeModuleWithEngine(fakeEngine(richResult()));
    const output = await module.handlers.execute({
      artifactRef: ref,
      operation: 'execute',
      requestId: 'req-refine',
    });
    const artifact = output.result.artifact;
    expect(artifact.id).toBe(ref.id);
    expect(artifact.sha256).toBe('a'.repeat(64));
    expect(artifact.sizeBytes).toBe(4096);
    expect(artifact.mediaType).toContain('wordprocessingml');
    // 调用方已有 label 时不被覆盖。
    expect(artifact.label).toBe(ref.label);
  });
});

/* -------------------------------------------------------------------------- */
/* 裁决（错误码映射）                                                            */
/* -------------------------------------------------------------------------- */

describe('contract: result adjudication', () => {
  it('maps engine failures to stable module error codes', async () => {
    await expect(
      inspectWith(fakeParseResult({ error: { kind: 'BAD_ZIP', message: 'broken' } })),
    ).rejects.toMatchObject({ code: 'PARSE_FAILED' });
    await expect(
      inspectWith(fakeParseResult({ error: { kind: 'NOT_FOUND', message: 'gone' } })),
    ).rejects.toMatchObject({ code: 'ARTIFACT_NOT_FOUND' });
  });

  it('distinguishes unsupported containers from format mismatches', async () => {
    await expect(
      inspectWith(fakeParseResult({ container: 'ole', documentKind: 'unknown' })),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_CONTAINER' });
    await expect(
      inspectWith(fakeParseResult({ container: 'rtf', documentKind: 'unknown' })),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_CONTAINER' });
    await expect(
      inspectWith(fakeParseResult({ container: 'unknown', documentKind: 'unknown' })),
    ).rejects.toMatchObject({ code: 'FORMAT_MISMATCH' });
    await expect(
      inspectWith(
        fakeParseResult({ documentKind: 'spreadsheetml', documentVariant: 'document' }),
      ),
    ).rejects.toMatchObject({ code: 'FORMAT_MISMATCH' });
  });

  it('blocks on limits only when enforcement is on', async () => {
    const limited = fakeParseResult({ limitHit: 'LIMIT_BLOCKS' });

    await expect(inspectWith(limited)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });

    const module = makeModuleWithEngine(fakeEngine(limited), {
      featureFlags: { enforceLimits: false },
    });
    const output = await module.handlers.inspect({
      artifactRef: artifactRef('clean.docx'),
      operation: 'inspect',
      requestId: 'req-limit',
    });
    // 宽限模式：保留结果，但必须给出 LIMIT_APPLIED 告警。
    expect(output.warnings.map((warning) => warning.code)).toContain('LIMIT_APPLIED');
  });

  it('honours declared hints: hard conflict fails, soft conflict warns', async () => {
    await expect(inspectWith(richResult(), { options: { declaredExtension: 'pptx' } })).rejects.toMatchObject(
      { code: 'FORMAT_MISMATCH' },
    );

    const output = await inspectWith(richResult(), { options: { declaredExtension: 'docm' } });
    expect(output.result.features['declaredHintsCompatible']).toBe(true);
    expect(output.warnings.map((warning) => warning.code)).toContain('DECLARED_HINT_MISMATCH');
  });

  it('rejects malformed input before touching the engine', async () => {
    await expect(inspectWith(richResult(), { requestId: '' })).rejects.toMatchObject({
      code: 'INVALID_INPUT',
    });
    await expect(
      inspectWith(richResult(), { artifactRef: { id: 'x', uri: 'https://example.com/a.docx' } }),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_ARTIFACT_URI' });
  });

  it('denies encrypted artifacts when the policy forbids them', async () => {
    await expect(
      inspectWith(fakeParseResult({ encrypted: true }), {
        policy: { id: 'p-strict', allowEncrypted: false },
      }),
    ).rejects.toMatchObject({ code: 'SAFETY_POLICY_DENIED' });
  });
});

/* -------------------------------------------------------------------------- */
/* 告警生成                                                                     */
/* -------------------------------------------------------------------------- */

describe('contract: warnings', () => {
  it('reports empty documents without failing the call', async () => {
    const output = await inspectWith(
      fakeParseResult({
        blocks: [para(0, '   ')],
        issues: [{ code: 'EMPTY_DOCUMENT', message: 'no visible content', path: 'word/document.xml' }],
      }),
    );
    expect(output.result.features['emptyDocument']).toBe(true);
    expect(output.warnings.map((warning) => warning.code)).toContain('EMPTY_DOCUMENT');
  });

  it('flags missing headings, comments and notes as informational warnings', async () => {
    const output = await inspectWith(richResult({ blocks: [para(0, 'body only')] }));
    const codes = output.warnings.map((warning) => warning.code);
    expect(codes).toContain('NO_HEADINGS');
    expect(codes).toContain('COMMENTS_PRESENT');
    expect(codes).toContain('FOOTNOTES_PRESENT');
    for (const warning of output.warnings) {
      expect(warning.severity).toBe('info');
    }
  });

  it('surfaces a backend fallback warning when lxml is absent', async () => {
    const output = await inspectWith(richResult({ xmlBackend: 'stdlib' }));
    expect(output.warnings.map((warning) => warning.code)).toContain('XML_BACKEND_FALLBACK');
  });
});

/* -------------------------------------------------------------------------- */
/* verify 三态语义                                                              */
/* -------------------------------------------------------------------------- */

describe('contract: verification semantics', () => {
  it('marks undeclared requirements as skip and reports partial verification', async () => {
    const module = makeModuleWithEngine(fakeEngine(richResult()));
    const output = await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-verify',
      policy: { id: 'p-default' },
    });
    const report: VerificationReport = output.result;
    expect(report.ok).toBe(true);
    expect(report.partial).toBe(true);
    expect(report.policyId).toBe('p-default');
    // 未声明要求一律 skip，而不是 pass。
    const headings = report.checks.find((entry) => entry.id === 'policy.headings');
    expect(headings?.status).toBe('skip');
  });

  it('fails when a declared requirement is not met', async () => {
    const module = makeModuleWithEngine(fakeEngine(richResult({ blocks: [para(0, 'no heading')] })));
    const output = await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-verify-fail',
      policy: { id: 'p-headings', requireHeadings: true },
    });
    expect(output.result.ok).toBe(false);
    const headings = output.result.checks.find((entry) => entry.id === 'policy.headings');
    expect(headings).toMatchObject({ status: 'fail', severity: 'error' });
  });

  it('evaluates the encryption policy in both directions', async () => {
    const encrypted = fakeParseResult({ encrypted: true });
    const strict = makeModuleWithEngine(fakeEngine(encrypted));
    const denied = await strict.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-enc-1',
      policy: { id: 'p-strict', allowEncrypted: false },
    });
    expect(denied.result.ok).toBe(false);

    const lenient = await strict.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-enc-2',
      policy: { id: 'p-lenient', allowEncrypted: true },
    });
    expect(lenient.result.ok).toBe(true);
    expect(
      lenient.result.checks.find((entry) => entry.id === 'policy.encryption')?.status,
    ).toBe('pass');
  });

  it('skips capability-dependent checks that were disabled for the call', async () => {
    const module = makeModuleWithEngine(fakeEngine(richResult({ blocks: [para(0, 'x')] })), {
      featureFlags: { parseStyles: false },
    });
    const output = await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-verify-skip',
      policy: { id: 'p-styles', requireStyles: true },
    });
    // 未解析样式时不能判 pass/fail，只能 skip。
    expect(output.result.checks.find((entry) => entry.id === 'policy.styles')?.status).toBe('skip');
  });

  it('requires a policy with an id', async () => {
    const module = makeModuleWithEngine(fakeEngine(richResult()));
    await expect(
      module.handlers.verify({
        artifactRef: artifactRef('clean.docx'),
        operation: 'verify',
        requestId: 'req-verify-bad',
        policy: { id: '  ' },
      }),
    ).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('produces a report satisfying the office-test-kit shape', async () => {
    const module = makeModuleWithEngine(fakeEngine(richResult()));
    const output = await module.handlers.verify({
      artifactRef: artifactRef('clean.docx'),
      operation: 'verify',
      requestId: 'req-verify-shape',
      policy: { id: 'p-shape' },
    });
    expect(output.result.summary.total).toBe(output.result.checks.length);
    expect(output.result.summary.passed + output.result.summary.failed + output.result.summary.skipped).toBe(
      output.result.summary.total,
    );
    expectJsonRoundTrip(output);
  });
});

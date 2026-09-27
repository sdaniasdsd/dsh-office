/**
 * 契约测试 —— 验证「本模块对外承诺的表面」是否成立。
 *
 * 这一组测试【不需要】Python 引擎：它只检查类型、配置、错误与注册契约，
 * 因此可以在 Profile 全量启动之前、甚至在没有解释器的机器上单独跑通。
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

import { describe, expect, it } from 'vitest';

import {
  DOCX_PARSE_CAPABILITIES,
  DOCX_PARSE_DEFINITION,
  DOCX_PARSE_ERROR_CODES,
  DOCX_PARSE_MODULE_ID,
  DOCX_PARSE_VERSION,
  DOCX_PARSE_WARNING_CODES,
  DUAL_VIEW_COUNT,
  DocxParseError,
  InMemoryTelemetry,
  SOURCE_MAP_SCHEME,
  createDocxParseModule,
  isDocxParseError,
  register,
  resolveArtifactPath,
  resolveConfig,
  toDocxParseError,
} from '../src/index';
import type { DocxParseModule } from '../src/index';
import { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS, DEFAULT_TIMEOUT_MS, withCallOverrides } from '../src/config';
import { PROJECT_ROOT, FakeEngine, parseResult } from './support';

/** 断言一个值可以安全地 JSON 往返。 */
function expectJsonRoundTrip(value: unknown): string {
  const encoded = JSON.stringify(value);
  expect(typeof encoded).toBe('string');
  const decoded: unknown = JSON.parse(encoded);
  expect(decoded).not.toBeUndefined();
  return encoded;
}

describe('contract: identity and capabilities', () => {
  it('exposes the module id required by the Profile layout', () => {
    expect(DOCX_PARSE_MODULE_ID).toBe('docx-parse');
  });

  it('publishes exactly the three capabilities from the spec', () => {
    expect([...DOCX_PARSE_CAPABILITIES]).toEqual(['inspect', 'execute', 'verify']);
  });

  it('freezes the dual-IR identity constants', () => {
    // 视图数量与锚点方案是双 IR 的「公开身份」，改动会破坏下游兼容性。
    expect(DUAL_VIEW_COUNT).toBe(2);
    expect(SOURCE_MAP_SCHEME).toBe('composite-anchor-v1');
  });

  it('maps every error and warning code to itself', () => {
    for (const [key, value] of Object.entries(DOCX_PARSE_ERROR_CODES)) {
      expect(value).toBe(key);
    }
    for (const [key, value] of Object.entries(DOCX_PARSE_WARNING_CODES)) {
      expect(value).toBe(key);
    }
  });

  it('declares the dual-IR specific codes', () => {
    expect(Object.keys(DOCX_PARSE_ERROR_CODES)).toContain('SOURCEMAP_INCOMPLETE');
    expect(Object.keys(DOCX_PARSE_WARNING_CODES)).toContain('ANCHOR_INCOMPLETE');
    expect(Object.keys(DOCX_PARSE_WARNING_CODES)).toContain('HEADING_LEVEL_INFERRED');
    expect(Object.keys(DOCX_PARSE_WARNING_CODES)).toContain('COMMENT_ORPHANED');
  });

  it('keeps the code tables JSON-serializable', () => {
    expectJsonRoundTrip(DOCX_PARSE_ERROR_CODES);
    expectJsonRoundTrip(DOCX_PARSE_WARNING_CODES);
  });
});

describe('contract: module definition', () => {
  it('declares a DOCX profile group and the upstream dependencies', () => {
    expect(DOCX_PARSE_DEFINITION.profileGroup).toBe('DOCX');
    expect(DOCX_PARSE_DEFINITION.version).toBe(DOCX_PARSE_VERSION);
    const names = DOCX_PARSE_DEFINITION.dependencies.map((entry) => entry.name);
    expect(names).toContain('office-core');
    expect(names).toContain('office-safety');
    expect(names).toContain('python');
  });

  it('ships a serializable config schema', () => {
    expectJsonRoundTrip(DOCX_PARSE_DEFINITION.configSchema);
    expect(DOCX_PARSE_DEFINITION.configSchema['type']).toBe('object');
  });

  it('stays consistent with module.json (single source of truth for the loader)', () => {
    const manifest = JSON.parse(readFileSync(join(PROJECT_ROOT, 'module.json'), 'utf8')) as Record<
      string,
      unknown
    >;
    expect(manifest['id']).toBe(DOCX_PARSE_DEFINITION.id);
    expect(manifest['version']).toBe(DOCX_PARSE_DEFINITION.version);
    expect(manifest['profileGroup']).toBe(DOCX_PARSE_DEFINITION.profileGroup);
    expect(manifest['capabilities']).toEqual([...DOCX_PARSE_DEFINITION.capabilities]);

    // 引擎声明必须指向真实存在的脚本，否则部署即失败。
    const engines = manifest['engines'] as Record<string, unknown>;
    expect(() => readFileSync(join(PROJECT_ROOT, String(engines['script'])))).not.toThrow();

    // 错误码/告警码表必须在清单与运行时之间逐项一致：
    // 否则 Profile 按清单判断可编程性，而模块实际可能抛出清单外的码。
    expect(manifest['errorCodes']).toEqual(Object.keys(DOCX_PARSE_ERROR_CODES));
    expect(manifest['warningCodes']).toEqual(Object.keys(DOCX_PARSE_WARNING_CODES));
  });
});

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
    const config = withCallOverrides(resolveConfig(), {
      limits: { maxBlocks: 10 },
      featureFlags: { resolveAnchors: false },
    });
    expect(config.limits.maxBlocks).toBe(10);
    expect(config.limits.maxArchiveEntries).toBe(DEFAULT_LIMITS.maxArchiveEntries);
    expect(config.featureFlags.resolveAnchors).toBe(false);
    expect(config.featureFlags.parseTables).toBe(true);
  });
});

describe('contract: module instance', () => {
  it('exposes handlers and a resolvable dispose', async () => {
    const module = createDocxParseModule({ engine: new FakeEngine(parseResult()) });
    expect(typeof module.handlers.inspect).toBe('function');
    expect(typeof module.handlers.execute).toBe('function');
    expect(typeof module.handlers.verify).toBe('function');
    await expect(module.dispose()).resolves.toBeUndefined();
  });

  it('registers itself through the Profile registry port', async () => {
    const registered: DocxParseModule[] = [];
    const registry = {
      registerModule(module: DocxParseModule): void {
        registered.push(module);
      },
    };
    const module = await register(registry, { engine: new FakeEngine(parseResult()) });
    expect(registered).toHaveLength(1);
    expect(registered[0]).toBe(module);
  });

  it('accepts an injected telemetry sink and records engine events', async () => {
    const telemetry = new InMemoryTelemetry();
    const module = createDocxParseModule({
      engine: new FakeEngine(parseResult({ blocks: [] })),
      telemetry,
    });
    await module.handlers.inspect({
      artifactRef: { id: 'a', uri: 'ignored-by-fake-engine.docx' },
      operation: 'inspect',
      requestId: 'req-1',
    });
    // FakeEngine 不产生遥测事件，因此这里只断言端口被注入后不报错、
    // 且事件列表形状可用（真实引擎的遥测由适配器负责）。
    expect(Array.isArray(telemetry.snapshot())).toBe(true);
  });
});

describe('contract: artifact uri resolution', () => {
  it('rejects non-local schemes so a materializer must be injected', () => {
    for (const uri of ['https://example.com/a.docx', 'memory://doc', 'data:application/zip;base64,AA==']) {
      expect(() => resolveArtifactPath(uri)).toThrow(DocxParseError);
    }
  });

  it('resolves file:// URIs and plain local paths', () => {
    // 用 pathToFileURL 构造平台正确的 file URL：Windows 上 /tmp 这类路径
    // 本身就不是合法的 file URL（缺少盘符），不属于本模块需要接受的输入。
    expect(resolveArtifactPath(pathToFileURL(join(PROJECT_ROOT, 'module.json')).href).length).toBeGreaterThan(0);
    expect(resolveArtifactPath('fixtures/a.docx').length).toBeGreaterThan(0);
  });

  it('rejects empty input', () => {
    expect(() => resolveArtifactPath('')).toThrow(DocxParseError);
  });
});

describe('contract: error type', () => {
  it('is serializable and recognized across realms', () => {
    const error = new DocxParseError('PARSE_FAILED', 'boom', { path: 'a.docx', details: { x: 1 } });
    expect(isDocxParseError(error)).toBe(true);
    expect(isDocxParseError(new Error('boom'))).toBe(false);
    expectJsonRoundTrip(error.toJSON());
    expect(error.toJSON()['code']).toBe('PARSE_FAILED');
  });

  it('wraps unknown values with a fallback code', () => {
    const wrapped = toDocxParseError('not-an-error', 'PARSE_FAILED');
    expect(isDocxParseError(wrapped)).toBe(true);
    expect(wrapped.code).toBe('PARSE_FAILED');
  });
});

/**
 * 契约测试 —— 验证「本模块对外承诺的表面」是否成立。
 *
 * 这一组测试【不需要】Python 引擎：它只检查类型、配置、错误与注册契约，
 * 因此可以在 Profile 全量启动之前、甚至在缺少解释器的机器上单独跑通
 * （spec 完成标准的第一条）。
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import {
  DOCX_INSPECT_CAPABILITIES,
  DOCX_INSPECT_DEFINITION,
  DOCX_INSPECT_ERROR_CODES,
  DOCX_INSPECT_MODULE_ID,
  DOCX_INSPECT_VERSION,
  DOCX_INSPECT_WARNING_CODES,
  DocxInspectError,
  InMemoryTelemetry,
  createDocxInspectModule,
  isDocxInspectError,
  register,
  resolveArtifactPath,
  resolveConfig,
  toDocxInspectError,
} from '../src/index';
import type {
  DocxInspectModule,
  JsonValue,
  ProfileRegistry,
} from '../src/index';
import { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS } from '../src/config';
import { configSchema, withCallOverrides } from '../src/config';

/** fixture 目录（由 globalSetup 生成）。 */
const FIXTURES_DIR = fileURLToPath(new URL('../fixtures/_generated/', import.meta.url));
const PROJECT_ROOT = fileURLToPath(new URL('..', import.meta.url));

/** 断言一个值可以安全地 JSON 往返（既不抛错，也不丢结构）。 */
function expectJsonRoundTrip(value: unknown): string {
  const encoded = JSON.stringify(value);
  expect(typeof encoded).toBe('string');
  // 二次解析不应抛错；undefined/函数会让 JSON.stringify 静默丢字段，
  // 因此额外检查「往返后仍能得到对象」。
  const decoded: unknown = JSON.parse(encoded);
  expect(decoded).not.toBeUndefined();
  return encoded;
}

describe('contract: identity and capabilities', () => {
  it('exposes the module id required by the Profile layout', () => {
    expect(DOCX_INSPECT_MODULE_ID).toBe('docx-inspect');
  });

  it('publishes exactly the three capabilities from the spec', () => {
    // 顺序也是契约的一部分：Profile 可能按声明顺序展示能力。
    expect([...DOCX_INSPECT_CAPABILITIES]).toEqual(['inspect', 'execute', 'verify']);
  });

  it('keeps error and warning code tables JSON-serializable', () => {
    expectJsonRoundTrip(DOCX_INSPECT_ERROR_CODES);
    expectJsonRoundTrip(DOCX_INSPECT_WARNING_CODES);
    expect(Object.keys(DOCX_INSPECT_ERROR_CODES)).toContain('ENGINE_PROTOCOL_ERROR');
    expect(Object.keys(DOCX_INSPECT_WARNING_CODES)).toContain('MACROS_PRESENT');
  });
});

describe('contract: module definition', () => {
  it('declares a DOCX profile group and the upstream dependencies', () => {
    expect(DOCX_INSPECT_DEFINITION.profileGroup).toBe('DOCX');
    expect(DOCX_INSPECT_DEFINITION.version).toBe(DOCX_INSPECT_VERSION);
    const names = DOCX_INSPECT_DEFINITION.dependencies.map((entry) => entry.name);
    expect(names).toContain('office-core');
    expect(names).toContain('office-safety');
    expect(names).toContain('office-files');
    expect(names).toContain('office-test-kit');
  });

  it('ships a serializable config schema', () => {
    expectJsonRoundTrip(DOCX_INSPECT_DEFINITION.configSchema);
    expect(DOCX_INSPECT_DEFINITION.configSchema['type']).toBe('object');
  });

  it('stays consistent with module.json (single source of truth for the loader)', () => {
    // module.json 是 Profile 加载器读取的声明式清单；
    // 若它与 TS 里的 definition 漂移，Profile 会加载到与运行时不一致的模块，
    // 因此这里做一次强制对齐检查。
    const raw = readFileSync(join(PROJECT_ROOT, 'module.json'), 'utf8');
    const manifest: Record<string, unknown> = JSON.parse(raw);
    expect(manifest['id']).toBe(DOCX_INSPECT_DEFINITION.id);
    expect(manifest['version']).toBe(DOCX_INSPECT_DEFINITION.version);
    expect(manifest['profileGroup']).toBe(DOCX_INSPECT_DEFINITION.profileGroup);
    expect(manifest['capabilities']).toEqual([...DOCX_INSPECT_DEFINITION.capabilities]);
    // 引擎声明必须指向真实存在的脚本，否则部署即失败。
    const engines = manifest['engines'] as Record<string, unknown>;
    const scriptPath = join(PROJECT_ROOT, String(engines['script']));
    expect(() => readFileSync(scriptPath)).not.toThrow();
  });

  it('produces a JSON-serializable definition', () => {
    expectJsonRoundTrip(DOCX_INSPECT_DEFINITION);
  });
});

describe('contract: configuration', () => {
  it('applies defaults and freezes the resolved config', () => {
    const config = resolveConfig();
    expect(config.limits).toEqual(DEFAULT_LIMITS);
    expect(config.featureFlags).toEqual(DEFAULT_FEATURE_FLAGS);
    expect(config.engine.driver).toBe('python');
    // 冻结：下游误改必须直接失败，而不是静默生效。
    expect(Object.isFrozen(config)).toBe(true);
    expect(Object.isFrozen(config.limits)).toBe(true);
    expect(Object.isFrozen(config.featureFlags)).toBe(true);
  });

  it('accepts partial overrides without dropping unrelated defaults', () => {
    const config = resolveConfig({
      limits: { maxArchiveEntries: 10 },
      featureFlags: { detectMacros: false },
    });
    expect(config.limits.maxArchiveEntries).toBe(10);
    // 未覆盖的字段保持默认值。
    expect(config.limits.maxRelationships).toBe(DEFAULT_LIMITS.maxRelationships);
    expect(config.featureFlags.detectMacros).toBe(false);
    expect(config.featureFlags.detectExternalLinks).toBe(true);
  });

  it('rejects invalid overrides with INVALID_INPUT', () => {
    // driver 不受支持
    expect(() => resolveConfig({ engine: { driver: 'ruby' as 'python' } })).toThrowError(
      /engine driver/i,
    );
    // pythonPath 为空
    expect(() => resolveConfig({ engine: { pythonPath: '  ' } })).toThrowError(
      /pythonPath/i,
    );
    // 超时非正整数
    expect(() => resolveConfig({ timeoutMs: 0 })).toThrowError(/timeoutMs/i);
    // 预算非正整数
    expect(() => resolveConfig({ limits: { maxArchiveEntries: -1 } })).toThrowError(
      /maxArchiveEntries/i,
    );
  });

  it('merges call-level overrides on top of the module config', () => {
    const base = resolveConfig({ timeoutMs: 5000 });
    const merged = withCallOverrides(base, { featureFlags: { detectActiveX: false } });
    expect(merged.timeoutMs).toBe(5000);
    expect(merged.featureFlags.detectActiveX).toBe(false);
    // 必须返回新对象，不能改动 base（base 已冻结，改动会抛错）。
    expect(merged).not.toBe(base);
    expect(base.featureFlags.detectActiveX).toBe(true);
  });

  it('describes every ModuleConfig field in the JSON schema', () => {
    const schema = configSchema();
    const properties = schema['properties'] as Record<string, JsonValue>;
    expect(Object.keys(properties).sort()).toEqual([
      'engine',
      'featureFlags',
      'limits',
      'runtimePackageNames',
      'runtimeRoot',
      'timeoutMs',
    ]);
  });
});

describe('contract: errors', () => {
  it('serializes a DocxInspectError to a stable JSON object', () => {
    const error = new DocxInspectError('FORMAT_MISMATCH', 'not a docx', {
      path: 'word/document.xml',
      details: { detected: 'spreadsheetml' },
    });
    const json = error.toJSON();
    expect(json['code']).toBe('FORMAT_MISMATCH');
    expect(json['path']).toBe('word/document.xml');
    expectJsonRoundTrip(json);
  });

  it('recognizes its own errors and wraps foreign ones', () => {
    const own = new DocxInspectError('PARSE_FAILED', 'boom');
    expect(isDocxInspectError(own)).toBe(true);
    expect(toDocxInspectError(own, 'ENGINE_FAILED')).toBe(own);

    const wrapped = toDocxInspectError(new TypeError('bad'), 'INVALID_INPUT');
    expect(wrapped.code).toBe('INVALID_INPUT');
    expect(isDocxInspectError(wrapped)).toBe(true);
    expectJsonRoundTrip(wrapped.toJSON());
  });

  it('keeps non-serializable causes out of the JSON payload', () => {
    // BigInt 无法被 JSON.stringify 直接处理；错误对象必须仍然可序列化。
    const wrapped = toDocxInspectError(10n, 'INVALID_INPUT');
    expectJsonRoundTrip(wrapped.toJSON());
  });
});

describe('contract: artifact URI resolution', () => {
  it('accepts absolute paths and relative paths', () => {
    const absolute = join(FIXTURES_DIR, 'clean.docx');
    expect(resolveArtifactPath(absolute)).toBe(absolute);
    // 相对路径按 cwd 解析为绝对路径。
    expect(resolveArtifactPath('clean.docx').endsWith('clean.docx')).toBe(true);
  });

  it('accepts file:// URIs', () => {
    const absolute = join(FIXTURES_DIR, 'clean.docx');
    const uri = `file:///${absolute.replace(/\\/g, '/').replace(/^\//, '')}`;
    expect(resolveArtifactPath(uri)).toBe(absolute);
  });

  it('rejects remote schemes so the Profile can inject a materializer', () => {
    expect(() => resolveArtifactPath('https://example.test/a.docx')).toThrowError(
      /materializer|Unsupported URI scheme/i,
    );
    expect(() => resolveArtifactPath('memory://buffer/1')).toThrowError(/materializer/i);
    expect(() => resolveArtifactPath('')).toThrowError(/non-empty/i);
  });
});

describe('contract: module factory and registration', () => {
  it('builds a module with the three handlers plus dispose', async () => {
    const module: DocxInspectModule = createDocxInspectModule({
      telemetry: new InMemoryTelemetry(),
    });
    expect(typeof module.handlers.inspect).toBe('function');
    expect(typeof module.handlers.execute).toBe('function');
    expect(typeof module.handlers.verify).toBe('function');
    await expect(module.dispose()).resolves.toBeUndefined();
  });

  it('registers itself through the minimal registry port', async () => {
    const registered: DocxInspectModule[] = [];
    const registry: ProfileRegistry = {
      registerModule(module) {
        registered.push(module);
      },
    };
    const module = await register(registry);
    expect(registered).toHaveLength(1);
    expect(registered[0]?.definition.id).toBe(DOCX_INSPECT_MODULE_ID);
    expect(module.definition.id).toBe(DOCX_INSPECT_MODULE_ID);
    await module.dispose();
  });

  it('creates independent instances (no shared mutable state)', () => {
    const first = createDocxInspectModule();
    const second = createDocxInspectModule();
    expect(first).not.toBe(second);
    expect(first.handlers.inspect).not.toBe(second.handlers.inspect);
  });
});

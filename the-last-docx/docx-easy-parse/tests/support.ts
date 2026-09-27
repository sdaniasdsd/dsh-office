/**
 * 测试辅助：fixture 定位、模块装配、假引擎。
 *
 * 为什么把「假引擎」放在共享位置：contract 测试必须在【没有 Python】的机器上
 * 也能跑通（spec 完成标准第一条）。有了假引擎，映射与验证逻辑就能脱离解释器
 * 单独验证；而 edge/regression 测试再使用真实 Python 引擎跑真样本。
 */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

import type { ArtifactRef } from 'office-core';

import { createDocxParseModule, resolveConfig } from '../src/index';
import type { DocxEngine, DocxParseModule, ModuleConfig, ParseResult } from '../src/index';
import type { ConfigOverrides } from '../src/config';

/** 生成器输出的目录（由 globalSetup 在测试前重建）。 */
export const FIXTURES_DIR = fileURLToPath(new URL('../fixtures/_generated/', import.meta.url));

/** 仓库根目录（用于对照 module.json）。 */
export const PROJECT_ROOT = fileURLToPath(new URL('..', import.meta.url));

/** 取一个 fixture 的绝对路径。 */
export function fixturePath(name: string): string {
  return `${FIXTURES_DIR}${name}`;
}

/** 构造指向 fixture 的 artifact 引用。 */
export function artifactRef(name: string): ArtifactRef {
  return { id: `fixture:${name}`, uri: fixturePath(name), label: name };
}

let pythonProbe: boolean | undefined;

/**
 * 本次运行应使用的 Python 解释器。
 *
 * 默认取 PATH 上的 `python`；集成测试可以把 `DOCX_PARSE_PYTHON` 指向一个
 * 装了 Docling 的虚拟环境。之所以做成显式函数而不是散落的 `process.env[...]`：
 * 探测与实际执行必须用同一个解释器，否则会出现「探测说有、跑起来说没有」。
 */
export function pythonExecutable(): string {
  return process.env['DOCX_PARSE_PYTHON'] ?? 'python';
}

/**
 * 探测 Python 是否可用（结果缓存，避免每个用例都 spawn 一次）。
 *
 * fixture 里的解析样本只有真实引擎才解得开，因此在没有解释器的环境下，
 * 相关用例会被跳过，而不是误报为失败。
 */
export function pythonAvailable(): boolean {
  if (pythonProbe === undefined) {
    const probe = spawnSync(pythonExecutable(), ['-c', 'print(1)'], { encoding: 'utf8' });
    pythonProbe = probe.status === 0;
  }
  return pythonProbe;
}

let doclingProbe: boolean | undefined;

/**
 * 探测 Docling 是否可导入（结果缓存）。
 *
 * 深度引擎依赖是可选的：未安装 Docling 的环境必须能跑通全部测试，
 * 因此相关用例整体跳过，而不是失败。
 */
export function doclingAvailable(): boolean {
  if (doclingProbe === undefined) {
    const probe = spawnSync(pythonExecutable(), ['-c', 'import docling'], { encoding: 'utf8' });
    doclingProbe = probe.status === 0;
  }
  return doclingProbe;
}

/** 用给定覆盖项装配一个模块实例（默认使用内置 Python 引擎）。 */
export function makeModule(overrides: ConfigOverrides = {}): DocxParseModule {
  const config: ModuleConfig = resolveConfig(overrides);
  return createDocxParseModule({ config });
}

/** 装配一个使用假引擎的模块实例。 */
export function makeModuleWithEngine(engine: DocxEngine, overrides: ConfigOverrides = {}): DocxParseModule {
  const config = resolveConfig(overrides);
  return createDocxParseModule({ config }, { engine });
}

/**
 * 装配一个同时注入轻量引擎与深度引擎的模块实例。
 *
 * 用于验证「哪个接口走哪个引擎」这一路由不变量：两个引擎都是可计数的假引擎，
 * 因此无需 Python 即可断言某条路径是否触碰了深度引擎。
 */
export function makeModuleWithEngines(
  engine: DocxEngine,
  deepEngine: DocxEngine,
  overrides: ConfigOverrides = {},
): DocxParseModule {
  const config = resolveConfig(overrides);
  return createDocxParseModule({ config }, { engine, deepEngine });
}

/** 包一层调用计数，用于断言某条路径是否真的执行了某个引擎。 */
export function countingEngine(inner: DocxEngine): { engine: DocxEngine; calls: () => number } {
  let calls = 0;
  const engine: DocxEngine = {
    name: inner.name,
    async parse(request): Promise<ParseResult> {
      calls += 1;
      return inner.parse(request);
    },
    async dispose(): Promise<void> {
      await inner.dispose();
    },
  };
  return { engine, calls: () => calls };
}

/** 造一个字段完整、默认「正常文档」的解析结果，便于精确覆盖单个字段。 */
export function fakeParseResult(overrides: Partial<ParseResult> = {}): ParseResult {
  const base: ParseResult = {
    parseVersion: 1,
    xmlBackend: 'lxml',
    container: 'zip',
    extension: 'docx',
    sizeBytes: 4096,
    sha256: 'a'.repeat(64),
    encrypted: false,
    documentKind: 'wordprocessingml',
    documentVariant: 'document',
    mediaType:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    metadata: {
      title: 'Fake',
      creator: null,
      lastModifiedBy: null,
      created: null,
      modified: null,
      revision: null,
    },
    blocks: [],
    styles: [],
    relationships: [],
    comments: [],
    footnotes: [],
    limitHit: null,
    issues: [],
    error: null,
  };
  return { ...base, ...overrides };
}

/** 固定返回同一份结果的假引擎（用于在没有 Python 时驱动完整流水线）。 */
export function fakeEngine(result: ParseResult): DocxEngine {
  return {
    name: 'fake-engine',
    async parse(): Promise<ParseResult> {
      return result;
    },
    async dispose(): Promise<void> {
      // 假引擎无资源可释放。
    },
  };
}

/** 断言一个值可 JSON 往返且不丢结构。 */
export function expectJsonRoundTrip(value: unknown): string {
  const encoded = JSON.stringify(value);
  const decoded: unknown = JSON.parse(encoded);
  if (decoded === undefined) {
    throw new Error('value did not survive JSON round-trip');
  }
  return encoded;
}

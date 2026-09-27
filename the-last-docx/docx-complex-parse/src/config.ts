/**
 * 引擎选择、资源预算、超时与能力开关。
 *
 * 职责边界：Profile 注入一份完整的 `ModuleConfig`；本文件只负责
 *   1. 提供默认值；
 *   2. 做「覆盖合并 + 合法性校验」；
 *   3. 冻结结果，防止下游误改。
 *
 * 本模块【从不】读取全局配置——这是 spec 明确要求的一条边界纪律，
 * 目的是让模块可以脱离 Profile 独立启动与测试。
 *
 * 默认值与 `module.json` 的 `configSchema` 必须逐项一致：两者是同一份契约的
 * 两种表达，出现分歧意味着「Profile 按 schema 校验通过的配置」与「模块实际
 * 使用的配置」不是一回事。
 */
import {
  COMPLEX_ENGINE_DRIVERS,
  type ComplexEngineDriver,
  type EngineConfig,
  type FeatureFlags,
  type JsonObject,
  type LimitConfig,
  type ModuleConfig,
} from './contract';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { resolveRuntime } from '@dsh-office-profile/docx-runtime';
import { DocxComplexParseError } from './errors';

/**
 * 默认资源预算。
 *
 * 取值依据：与 docx-parse / docx-inspect 保持同一口径——三个模块面对的是同一类
 * 容器（ZIP + OOXML），预算分歧会让同一个文件在不同模块里一个能读、一个读不了。
 * 单部件 32 MiB / 总量 256 MiB 能在遇到解压炸弹时及时止损。
 */
export const DEFAULT_LIMITS: LimitConfig = {
  maxArchiveEntries: 4096,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxRelationships: 4096,
  maxBlocks: 50_000,
  maxTableCells: 200_000,
  maxFloatingObjects: 20_000,
  maxPages: 5_000,
};

/**
 * 默认能力开关：本模块五项处理范围全部开启。
 *
 * 唯一的例外是 `requirePageGeometry`——把「拿不到页几何」当成硬失败是调用方的
 * 策略选择，不是模块的默认立场。默认降级（结构坐标仍然有效）比默认拒绝更可用。
 */
export const DEFAULT_FEATURE_FLAGS: FeatureFlags = {
  parseTables: true,
  parseMerges: true,
  parsePageBreaks: true,
  parseSections: true,
  parseFloatingObjects: true,
  parsePageGeometry: true,
  computeConfidence: true,
  requirePageGeometry: false,
  enforceLimits: true,
};

/** 默认超时：60 秒。版面模型比纯 XML 读取重得多，预算相应放大。 */
export const DEFAULT_TIMEOUT_MS = 60_000;

/**
 * 默认置信度下限。
 *
 * 低于此值的结论会附带 `LOW_CONFIDENCE` 告警，但**不会**被剔除——「不够可信」
 * 是给调用方的信息，不是替调用方做的决定。
 */
export const DEFAULT_CONFIDENCE_FLOOR = 0.5;

/** 默认引擎配置：使用 PATH 上的 `python` 驱动内置的 rdocx 引擎。 */
export const DEFAULT_ENGINE: EngineConfig = {
  driver: 'rdocx',
  pythonPath: 'python',
};

/**
 * 深冻结一份配置。
 *
 * 冻结 `engine` / `limits` / `featureFlags` 三个子对象本身，防止调用方在拿到
 * `ModuleConfig` 后误改，导致同一次请求内行为不一致。
 */
function freezeConfig(config: ModuleConfig): ModuleConfig {
  return Object.freeze({
    engine: Object.freeze({ ...config.engine }),
    runtimeRoot: config.runtimeRoot,
    runtimePackageNames: config.runtimePackageNames === undefined ? undefined : Object.freeze([...config.runtimePackageNames]),
    limits: Object.freeze({ ...config.limits }),
    timeoutMs: config.timeoutMs,
    featureFlags: Object.freeze({ ...config.featureFlags }),
    confidenceFloor: config.confidenceFloor,
  });
}

/** 配置覆盖项：允许 Profile 只提供它关心的增量。 */
export interface ConfigOverrides {
  engine?: Partial<EngineConfig>;
  runtimeRoot?: string;
  runtimePackageNames?: readonly string[];
  limits?: Partial<LimitConfig>;
  timeoutMs?: number;
  featureFlags?: Partial<FeatureFlags>;
  confidenceFloor?: number;
}

/** 校验一个数值型预算：必须是正整数。 */
function requireBudget(name: string, value: number): number {
  if (!Number.isInteger(value) || value <= 0) {
    throw new DocxComplexParseError('INVALID_INPUT', `Invalid limit "${name}"`, {
      details: { name, value },
    });
  }
  return value;
}

/** 逐项校验资源预算。 */
function validateLimits(limits: LimitConfig): LimitConfig {
  return {
    maxArchiveEntries: requireBudget('maxArchiveEntries', limits.maxArchiveEntries),
    maxEntryUncompressedBytes: requireBudget(
      'maxEntryUncompressedBytes',
      limits.maxEntryUncompressedBytes,
    ),
    maxTotalUncompressedBytes: requireBudget(
      'maxTotalUncompressedBytes',
      limits.maxTotalUncompressedBytes,
    ),
    maxRelationships: requireBudget('maxRelationships', limits.maxRelationships),
    maxBlocks: requireBudget('maxBlocks', limits.maxBlocks),
    maxTableCells: requireBudget('maxTableCells', limits.maxTableCells),
    maxFloatingObjects: requireBudget('maxFloatingObjects', limits.maxFloatingObjects),
    maxPages: requireBudget('maxPages', limits.maxPages),
  };
}

function isEngineDriver(value: string): value is ComplexEngineDriver {
  return (COMPLEX_ENGINE_DRIVERS as readonly string[]).includes(value);
}

/**
 * 把 Profile 注入的值合并到模块默认值之上，并返回冻结后的完整配置。
 *
 * 注意这里**不**校验 driver 是否已实现：`docling` / `mammoth` / `markitdown` 是
 * 合法配置，只是当前没有内置实现。配置合法而引擎缺失，应该在调用时以
 * `ENGINE_UNAVAILABLE` 明确失败，而不是在配置阶段悄悄拒绝——否则宿主会以为
 * 自己写错了配置。
 *
 * @throws DocxComplexParseError(INVALID_INPUT) driver 不在引擎族内、
 *         pythonPath 为空、或 timeoutMs/limits/confidenceFloor 非法时。
 */
export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  const engine: EngineConfig = { ...DEFAULT_ENGINE, ...overrides.engine };
  const callerPath = fileURLToPath(import.meta.url);
  const runtime = resolveRuntime({
    runtimeRoot: overrides.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames,
    callerPath,
    bundledRuntimeRoot: join(dirname(callerPath), '..', 'runtime', 'win32-x64'),
    components: [{ id: 'python', envVar: 'DOCX_PYTHON', defaultEntry: 'python/python.exe' }],
  });
  if (overrides.engine?.pythonPath === undefined && runtime.components.python?.path !== undefined) {
    engine.pythonPath = runtime.components.python.path;
  }
  if (!isEngineDriver(engine.driver)) {
    throw new DocxComplexParseError('INVALID_INPUT', 'Unsupported engine driver', {
      details: { driver: String(engine.driver), supported: [...COMPLEX_ENGINE_DRIVERS] },
    });
  }
  if (typeof engine.pythonPath !== 'string' || engine.pythonPath.trim() === '') {
    throw new DocxComplexParseError('INVALID_INPUT', 'engine.pythonPath must be a non-empty string');
  }
  if (engine.scriptPath !== undefined && engine.scriptPath.trim() === '') {
    throw new DocxComplexParseError('INVALID_INPUT', 'engine.scriptPath must not be empty');
  }

  const timeoutMs = overrides.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  if (!Number.isInteger(timeoutMs) || timeoutMs <= 0) {
    throw new DocxComplexParseError('INVALID_INPUT', 'timeoutMs must be a positive integer', {
      details: { timeoutMs },
    });
  }

  const confidenceFloor = overrides.confidenceFloor ?? DEFAULT_CONFIDENCE_FLOOR;
  if (
    typeof confidenceFloor !== 'number' ||
    !Number.isFinite(confidenceFloor) ||
    confidenceFloor < 0 ||
    confidenceFloor > 1
  ) {
    throw new DocxComplexParseError(
      'INVALID_INPUT',
      'confidenceFloor must be a number within 0..1',
      { details: { confidenceFloor } },
    );
  }

  return freezeConfig({
    engine,
    runtimeRoot: overrides.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames,
    limits: validateLimits({ ...DEFAULT_LIMITS, ...overrides.limits }),
    timeoutMs,
    featureFlags: { ...DEFAULT_FEATURE_FLAGS, ...overrides.featureFlags },
    confidenceFloor,
  });
}

/**
 * 在已有配置之上应用「单次调用级」覆盖。
 *
 * 场景：某次调用需要更小的超时或更严的预算，但不应影响模块的全局配置。
 */
export function withCallOverrides(base: ModuleConfig, overrides: ConfigOverrides = {}): ModuleConfig {
  const merged: ConfigOverrides = {
    engine: { ...base.engine, ...overrides.engine },
    runtimeRoot: overrides.runtimeRoot ?? base.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames ?? base.runtimePackageNames,
    limits: { ...base.limits, ...overrides.limits },
    featureFlags: { ...base.featureFlags, ...overrides.featureFlags },
    timeoutMs: overrides.timeoutMs ?? base.timeoutMs,
    confidenceFloor: overrides.confidenceFloor ?? base.confidenceFloor,
  };
  return resolveConfig(merged);
}

/** 把 `DocxComplexParseOptions` 里的覆盖项转成 `ConfigOverrides`。 */
export function toConfigOverrides(options: {
  limits?: Partial<LimitConfig>;
  featureFlags?: Partial<FeatureFlags>;
  timeoutMs?: number;
  confidenceFloor?: number;
}): ConfigOverrides {
  const overrides: ConfigOverrides = {};
  if (options.limits !== undefined) overrides.limits = options.limits;
  if (options.featureFlags !== undefined) overrides.featureFlags = options.featureFlags;
  if (options.timeoutMs !== undefined) overrides.timeoutMs = options.timeoutMs;
  if (options.confidenceFloor !== undefined) overrides.confidenceFloor = options.confidenceFloor;
  return overrides;
}

/**
 * 描述 `ModuleConfig` 的 JSON-Schema 子集。
 *
 * 与 `module.json` 中的 `configSchema` 保持一致：Profile 用它来校验注入的配置，
 * 从而在模块启动前就发现配置错误（fail fast）。
 */
export function configSchema(): JsonObject {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      runtimeRoot: { type: 'string' },
      runtimePackageNames: { type: 'array', items: { type: 'string' } },
      engine: {
        type: 'object',
        additionalProperties: false,
        properties: {
          driver: {
            type: 'string',
            enum: [...COMPLEX_ENGINE_DRIVERS],
            default: DEFAULT_ENGINE.driver,
          },
          pythonPath: { type: 'string', default: DEFAULT_ENGINE.pythonPath },
          scriptPath: { type: 'string' },
          env: { type: 'object', additionalProperties: { type: 'string' } },
        },
      },
      limits: {
        type: 'object',
        additionalProperties: false,
        properties: {
          maxArchiveEntries: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxArchiveEntries,
          },
          maxEntryUncompressedBytes: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxEntryUncompressedBytes,
          },
          maxTotalUncompressedBytes: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxTotalUncompressedBytes,
          },
          maxRelationships: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxRelationships,
          },
          maxBlocks: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxBlocks },
          maxTableCells: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxTableCells },
          maxFloatingObjects: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxFloatingObjects,
          },
          maxPages: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxPages },
        },
      },
      timeoutMs: { type: 'integer', minimum: 1, default: DEFAULT_TIMEOUT_MS },
      confidenceFloor: {
        type: 'number',
        minimum: 0,
        maximum: 1,
        default: DEFAULT_CONFIDENCE_FLOOR,
      },
      featureFlags: {
        type: 'object',
        additionalProperties: false,
        properties: {
          parseTables: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseTables },
          parseMerges: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseMerges },
          parsePageBreaks: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parsePageBreaks },
          parseSections: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseSections },
          parseFloatingObjects: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.parseFloatingObjects,
          },
          parsePageGeometry: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parsePageGeometry },
          computeConfidence: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.computeConfidence },
          requirePageGeometry: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.requirePageGeometry,
          },
          enforceLimits: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.enforceLimits },
        },
      },
    },
  };
}

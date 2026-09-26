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
 */
import {
  type EngineConfig,
  type FeatureFlags,
  type JsonObject,
  type LimitConfig,
  type ModuleConfig,
} from './contract';
import { DocxParseError } from './errors';

/**
 * 默认资源预算。
 *
 * 取值依据：真实 DOCX 的正文通常有几十到几千个块。这里给到 50000 块的上限，
 * 既能容纳极长的报告/论文，又能在遇到恶意构造时及时止损。
 * 单部件 32 MiB / 总量 256 MiB 与 docx-inspect 保持一致：
 * 两个模块面对的是同一类容器，预算口径不应出现分歧。
 */
export const DEFAULT_LIMITS: LimitConfig = {
  maxArchiveEntries: 4096,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxRelationships: 4096,
  maxBlocks: 50_000,
  maxTableCells: 200_000,
  maxAnnotations: 10_000,
};

/** 默认能力开关：解析类能力全部开启，即「完整双 IR」。 */
export const DEFAULT_FEATURE_FLAGS: FeatureFlags = {
  parseHeadings: true,
  parseTables: true,
  parseStyles: true,
  parseComments: true,
  parseFootnotes: true,
  // 身份锚点是双 IR 的意义所在，默认开启。
  resolveAnchors: true,
  enforceLimits: true,
};

/** 默认超时：30 秒。解析比探测更重，预算相应放大。 */
export const DEFAULT_TIMEOUT_MS = 30_000;

/** 默认引擎配置：使用 PATH 上的 `python`。 */
export const DEFAULT_ENGINE: EngineConfig = {
  driver: 'python',
  pythonPath: 'python',
};

/**
 * 深冻结一份配置。
 *
 * 冻结的是 `limits` / `featureFlags` / `engine` 三个子对象本身，
 * 防止调用方在拿到 `ModuleConfig` 后误改，导致同一次请求内行为不一致。
 */
function freezeConfig(config: ModuleConfig): ModuleConfig {
  const frozen: ModuleConfig = {
    engine: Object.freeze({ ...config.engine }),
    limits: Object.freeze({ ...config.limits }),
    timeoutMs: config.timeoutMs,
    featureFlags: Object.freeze({ ...config.featureFlags }),
  };
  return Object.freeze(frozen);
}

/** 配置覆盖项：允许 Profile 只提供它关心的增量。 */
export interface ConfigOverrides {
  engine?: Partial<EngineConfig>;
  limits?: Partial<LimitConfig>;
  timeoutMs?: number;
  featureFlags?: Partial<FeatureFlags>;
}

/** 校验一个数值型预算：必须是正整数。 */
function requireBudget(name: string, value: number): number {
  if (!Number.isInteger(value) || value <= 0) {
    throw new DocxParseError('INVALID_INPUT', `Invalid limit "${name}"`, {
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
    maxAnnotations: requireBudget('maxAnnotations', limits.maxAnnotations),
  };
}

/**
 * 把 Profile 注入的值合并到模块默认值之上，并返回冻结后的完整配置。
 *
 * @throws DocxParseError(INVALID_INPUT) 当 driver 不受支持、pythonPath 为空
 *         或 timeoutMs/limits 非法时。
 */
export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  const engine: EngineConfig = {
    ...DEFAULT_ENGINE,
    ...overrides.engine,
  };
  if (engine.driver !== 'python') {
    throw new DocxParseError('INVALID_INPUT', 'Unsupported engine driver', {
      details: { driver: String(engine.driver), supported: ['python'] },
    });
  }
  if (typeof engine.pythonPath !== 'string' || engine.pythonPath.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'engine.pythonPath must be a non-empty string');
  }

  const timeoutMs = overrides.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  if (!Number.isInteger(timeoutMs) || timeoutMs <= 0) {
    throw new DocxParseError('INVALID_INPUT', 'timeoutMs must be a positive integer', {
      details: { timeoutMs },
    });
  }

  return freezeConfig({
    engine,
    limits: validateLimits({ ...DEFAULT_LIMITS, ...overrides.limits }),
    timeoutMs,
    featureFlags: { ...DEFAULT_FEATURE_FLAGS, ...overrides.featureFlags },
  });
}

/**
 * 在已有配置之上应用「单次调用级」覆盖。
 *
 * 场景：某次调用需要更小的超时或更严的预算，但不应影响模块的全局配置。
 */
export function withCallOverrides(
  base: ModuleConfig,
  overrides: ConfigOverrides = {},
): ModuleConfig {
  const merged: ConfigOverrides = {
    engine: { ...base.engine, ...overrides.engine },
    limits: { ...base.limits, ...overrides.limits },
    featureFlags: { ...base.featureFlags, ...overrides.featureFlags },
  };
  if (overrides.timeoutMs !== undefined) {
    merged.timeoutMs = overrides.timeoutMs;
  } else {
    merged.timeoutMs = base.timeoutMs;
  }
  return resolveConfig(merged);
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
      engine: {
        type: 'object',
        additionalProperties: false,
        properties: {
          driver: { type: 'string', enum: ['python'], default: DEFAULT_ENGINE.driver },
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
          maxAnnotations: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxAnnotations,
          },
        },
      },
      timeoutMs: { type: 'integer', minimum: 1, default: DEFAULT_TIMEOUT_MS },
      featureFlags: {
        type: 'object',
        additionalProperties: false,
        properties: {
          parseHeadings: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseHeadings },
          parseTables: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseTables },
          parseStyles: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseStyles },
          parseComments: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseComments },
          parseFootnotes: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseFootnotes },
          resolveAnchors: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.resolveAnchors },
          enforceLimits: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.enforceLimits },
        },
      },
    },
  };
}

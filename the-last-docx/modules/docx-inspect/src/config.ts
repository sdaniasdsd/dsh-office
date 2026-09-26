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
import { DocxInspectError } from './errors';

/**
 * 默认资源预算。
 *
 * 取值依据：真实 DOCX 的 `document.xml` 通常在数十 KB 到数 MB 量级，
 * 32 MiB 的单部件上限足以覆盖超大报告，同时对解压炸弹立即触发拦截。
 * 总预算 256 MiB 同理：真实文档远低于此，恶意包会迅速越界。
 */
export const DEFAULT_LIMITS: LimitConfig = {
  maxArchiveEntries: 4096,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxRelationships: 4096,
  maxExternalTargets: 512,
};

/** 默认能力开关：全部开启，即「完整探测」。 */
export const DEFAULT_FEATURE_FLAGS: FeatureFlags = {
  detectMacros: true,
  detectExternalLinks: true,
  detectEmbeddedObjects: true,
  detectAltChunks: true,
  detectActiveX: true,
  detectDdeFields: true,
  // 宏语义分析是【可选增强】（依赖 Python 侧的 oletools），默认关闭：
  // 未安装 oletools 的环境不会因它产生噪音，需要时显式开启即可。
  analyzeMacroCode: false,
  enforceLimits: true,
};

/** 默认超时：20 秒，足够覆盖大文档，也能及时掐死僵死进程。 */
export const DEFAULT_TIMEOUT_MS = 20_000;

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

/**
 * 校验一个数值型预算：必须是正整数且不超过给定上限。
 *
 * 之所以上限参数存在但传 `MAX_SAFE_INTEGER`：保留将来对某项预算单独设硬顶的能力，
 * 而不必为此改写调用点。
 */
function requireBudget(name: string, value: number, ceiling: number): number {
  if (!Number.isInteger(value) || value <= 0 || value > ceiling) {
    throw new DocxInspectError('INVALID_INPUT', `Invalid limit "${name}"`, {
      details: { name, value, ceiling },
    });
  }
  return value;
}

/** 逐项校验资源预算。 */
function validateLimits(limits: LimitConfig): LimitConfig {
  return {
    maxArchiveEntries: requireBudget(
      'maxArchiveEntries',
      limits.maxArchiveEntries,
      Number.MAX_SAFE_INTEGER,
    ),
    maxEntryUncompressedBytes: requireBudget(
      'maxEntryUncompressedBytes',
      limits.maxEntryUncompressedBytes,
      Number.MAX_SAFE_INTEGER,
    ),
    maxTotalUncompressedBytes: requireBudget(
      'maxTotalUncompressedBytes',
      limits.maxTotalUncompressedBytes,
      Number.MAX_SAFE_INTEGER,
    ),
    maxRelationships: requireBudget(
      'maxRelationships',
      limits.maxRelationships,
      Number.MAX_SAFE_INTEGER,
    ),
    maxExternalTargets: requireBudget(
      'maxExternalTargets',
      limits.maxExternalTargets,
      Number.MAX_SAFE_INTEGER,
    ),
  };
}

/**
 * 把 Profile 注入的值合并到模块默认值之上，并返回冻结后的完整配置。
 *
 * 接受 Partial 的好处：Profile 可以只提供差异项，让 `module.json` 的
 * `configSchema` 全部字段保持可选，降低接入成本。
 *
 * @throws DocxInspectError(INVALID_INPUT) 当 driver 不受支持、pythonPath 为空
 *         或 timeoutMs/limits 非法时。
 */
export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  const engine: EngineConfig = {
    ...DEFAULT_ENGINE,
    ...overrides.engine,
  };
  if (engine.driver !== 'python') {
    throw new DocxInspectError('INVALID_INPUT', `Unsupported engine driver`, {
      details: { driver: String(engine.driver), supported: ['python'] },
    });
  }
  if (typeof engine.pythonPath !== 'string' || engine.pythonPath.trim() === '') {
    throw new DocxInspectError('INVALID_INPUT', 'engine.pythonPath must be a non-empty string');
  }

  const timeoutMs = overrides.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  if (!Number.isInteger(timeoutMs) || timeoutMs <= 0) {
    throw new DocxInspectError('INVALID_INPUT', 'timeoutMs must be a positive integer', {
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
 * 因此这里以 `base` 为底，仅覆盖显式传入的字段。
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
          maxExternalTargets: {
            type: 'integer',
            minimum: 1,
            default: DEFAULT_LIMITS.maxExternalTargets,
          },
        },
      },
      timeoutMs: { type: 'integer', minimum: 1, default: DEFAULT_TIMEOUT_MS },
      featureFlags: {
        type: 'object',
        additionalProperties: false,
        properties: {
          detectMacros: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.detectMacros },
          detectExternalLinks: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.detectExternalLinks,
          },
          detectEmbeddedObjects: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.detectEmbeddedObjects,
          },
          detectAltChunks: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.detectAltChunks },
          detectActiveX: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.detectActiveX },
          detectDdeFields: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.detectDdeFields },
          analyzeMacroCode: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.analyzeMacroCode,
          },
          enforceLimits: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.enforceLimits },
        },
      },
    },
  };
}

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
  type DeepEngineConfig,
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
 * 取值依据：解析会把正文块、表格单元格、样式、批注、脚注全部展开为结构化记录，
 * 因此除了归档层预算（防解压炸弹）之外，还需要对「展开后的规模」单独设限，
 * 否则一个正常大小的包也可能产出百万级记录拖垮上层。
 */
export const DEFAULT_LIMITS: LimitConfig = {
  maxArchiveEntries: 4096,
  maxEntryUncompressedBytes: 32 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxRelationships: 4096,
  maxBlocks: 20000,
  maxTableCells: 20000,
  maxStyles: 2000,
  maxComments: 2000,
  maxFootnotes: 2000,
};

/** 默认能力开关：全部开启，即「完整解析」。 */
export const DEFAULT_FEATURE_FLAGS: FeatureFlags = {
  parseParagraphs: true,
  parseHeadings: true,
  parseTables: true,
  parseStyles: true,
  parseRelationships: true,
  parseComments: true,
  parseFootnotes: true,
  extractMetadata: true,
  enforceLimits: true,
};

/** 默认超时：20 秒，足够覆盖大文档，也能及时掐死僵死进程。 */
export const DEFAULT_TIMEOUT_MS = 20_000;

/**
 * 深度引擎的默认超时：120 秒。
 *
 * 比轻量引擎宽一个量级，因为深度引擎每次调用都要冷启动模型；
 * 20 秒的预算会把「模型加载慢」直接变成 ENGINE_TIMEOUT 误判。
 */
export const DEFAULT_DEEP_TIMEOUT_MS = 120_000;

/** 默认引擎配置：使用 PATH 上的 `python`，不启用深度引擎。 */
export const DEFAULT_ENGINE: EngineConfig = {
  driver: 'python',
  pythonPath: 'python',
};

/**
 * 深冻结一份配置。
 *
 * 冻结的是 `limits` / `featureFlags` / `engine`（含嵌套的 `deep`）三个子对象，
 * 防止调用方在拿到 `ModuleConfig` 后误改，导致同一次请求内行为不一致。
 */
function freezeConfig(config: ModuleConfig): ModuleConfig {
  const engine: EngineConfig = { ...config.engine };
  if (engine.deep !== undefined) {
    engine.deep = Object.freeze({ ...engine.deep });
  }
  const frozen: ModuleConfig = {
    engine: Object.freeze(engine),
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
    throw new DocxParseError('INVALID_INPUT', `Invalid limit "${name}"`, {
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
    maxBlocks: requireBudget('maxBlocks', limits.maxBlocks, Number.MAX_SAFE_INTEGER),
    maxTableCells: requireBudget('maxTableCells', limits.maxTableCells, Number.MAX_SAFE_INTEGER),
    maxStyles: requireBudget('maxStyles', limits.maxStyles, Number.MAX_SAFE_INTEGER),
    maxComments: requireBudget('maxComments', limits.maxComments, Number.MAX_SAFE_INTEGER),
    maxFootnotes: requireBudget('maxFootnotes', limits.maxFootnotes, Number.MAX_SAFE_INTEGER),
  };
}

/**
 * 校验并归一化深度引擎配置。
 *
 * `pythonPath` / `scriptPath` 未提供时保持 undefined，由适配器在运行时
 * 回退到 `EngineConfig.pythonPath` 与内置脚本路径——模块层不猜测文件系统布局。
 *
 * @throws DocxParseError(INVALID_INPUT) driver 不受支持或 timeoutMs 非法
 */
function validateDeepEngine(deep: DeepEngineConfig): DeepEngineConfig {
  if (deep.driver !== 'docling') {
    throw new DocxParseError('INVALID_INPUT', 'Unsupported deep engine driver', {
      details: { driver: String(deep.driver), supported: ['docling'] },
    });
  }
  if (deep.timeoutMs !== undefined && (!Number.isInteger(deep.timeoutMs) || deep.timeoutMs <= 0)) {
    throw new DocxParseError('INVALID_INPUT', 'engine.deep.timeoutMs must be a positive integer', {
      details: { timeoutMs: deep.timeoutMs },
    });
  }
  const normalized: DeepEngineConfig = { driver: deep.driver };
  if (deep.pythonPath !== undefined) normalized.pythonPath = deep.pythonPath;
  if (deep.scriptPath !== undefined) normalized.scriptPath = deep.scriptPath;
  if (deep.env !== undefined) normalized.env = { ...deep.env };
  if (deep.timeoutMs !== undefined) normalized.timeoutMs = deep.timeoutMs;
  return normalized;
}

/**
 * 把 Profile 注入的值合并到模块默认值之上，并返回冻结后的完整配置。
 *
 * 接受 Partial 的好处：Profile 可以只提供差异项，让 `module.json` 的
 * `configSchema` 全部字段保持可选，降低接入成本。
 *
 * @throws DocxParseError(INVALID_INPUT) 当 driver 不受支持、pythonPath 为空、
 *         deep 配置非法或 timeoutMs/limits 非法时。
 */
export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  const engine: EngineConfig = {
    ...DEFAULT_ENGINE,
    ...overrides.engine,
  };
  if (engine.driver !== 'python') {
    throw new DocxParseError('INVALID_INPUT', `Unsupported engine driver`, {
      details: { driver: String(engine.driver), supported: ['python'] },
    });
  }
  if (typeof engine.pythonPath !== 'string' || engine.pythonPath.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'engine.pythonPath must be a non-empty string');
  }
  // 校验并归一化深度引擎配置（存在时才校验）。
  if (engine.deep !== undefined) {
    engine.deep = validateDeepEngine({ ...engine.deep });
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
          // 深度引擎：仅作用于 execute / verify；inspect 恒用轻量引擎。
          deep: {
            type: 'object',
            additionalProperties: false,
            properties: {
              driver: { type: 'string', enum: ['docling'] },
              pythonPath: { type: 'string' },
              scriptPath: { type: 'string' },
              env: { type: 'object', additionalProperties: { type: 'string' } },
              timeoutMs: { type: 'integer', minimum: 1, default: DEFAULT_DEEP_TIMEOUT_MS },
            },
            required: ['driver'],
          },
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
          maxStyles: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxStyles },
          maxComments: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxComments },
          maxFootnotes: { type: 'integer', minimum: 1, default: DEFAULT_LIMITS.maxFootnotes },
        },
      },
      timeoutMs: { type: 'integer', minimum: 1, default: DEFAULT_TIMEOUT_MS },
      featureFlags: {
        type: 'object',
        additionalProperties: false,
        properties: {
          parseParagraphs: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.parseParagraphs,
          },
          parseHeadings: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseHeadings },
          parseTables: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseTables },
          parseStyles: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseStyles },
          parseRelationships: {
            type: 'boolean',
            default: DEFAULT_FEATURE_FLAGS.parseRelationships,
          },
          parseComments: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseComments },
          parseFootnotes: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.parseFootnotes },
          extractMetadata: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.extractMetadata },
          enforceLimits: { type: 'boolean', default: DEFAULT_FEATURE_FLAGS.enforceLimits },
        },
      },
    },
  };
}

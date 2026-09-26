import type { ConfigOverrides, ModuleConfig, StyleFeatureFlags, StyleLimits } from './contract';

export const DEFAULT_LIMITS: StyleLimits = {
  maxInputBytes: 64 * 1024 * 1024,
  maxOutputBytes: 128 * 1024 * 1024,
  maxStyles: 500,
};

export const DEFAULT_FEATURE_FLAGS: StyleFeatureFlags = {
  /** Authoring named styles. Turning it off leaves inspect/verify working. */
  allowStyleDefinitions: true,
  /**
   * Rewriting a definition the document already carries changes the appearance
   * of every paragraph that uses that style, so it is a separate decision from
   * adding a style the document lacks.
   */
  allowRedefineExisting: true,
};

export const DEFAULT_TIMEOUT_MS = 30_000;

export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  return {
    engine: { driver: 'docx4j-core-ts' },
    limits: { ...DEFAULT_LIMITS, ...overrides.limits },
    timeoutMs: overrides.timeoutMs ?? DEFAULT_TIMEOUT_MS,
    featureFlags: { ...DEFAULT_FEATURE_FLAGS, ...overrides.featureFlags },
  };
}

export function withCallOverrides(config: ModuleConfig, overrides?: ConfigOverrides): ModuleConfig {
  if (!overrides) return config;
  return {
    ...config,
    limits: { ...config.limits, ...overrides.limits },
    timeoutMs: overrides.timeoutMs ?? config.timeoutMs,
    featureFlags: { ...config.featureFlags, ...overrides.featureFlags },
  };
}

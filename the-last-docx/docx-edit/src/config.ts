import type { ConfigOverrides, EditFeatureFlags, EditLimits, ModuleConfig } from './contract';

export const DEFAULT_LIMITS: EditLimits = {
  maxInputBytes: 64 * 1024 * 1024,
  maxOutputBytes: 128 * 1024 * 1024,
  maxOperations: 1000,
};

export const DEFAULT_FEATURE_FLAGS: EditFeatureFlags = {
  allowTextReplacement: true,
  allowComments: true,
  allowTrackedChanges: true,
  allowAcceptReject: false,
  requireExpectedDigest: true,
  allowFormatting: true,
  allowTableFormatting: true,
  allowBlockEdits: true,
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

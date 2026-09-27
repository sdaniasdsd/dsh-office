import type { ConfigOverrides, ModuleConfig, RenderFeatureFlags, RenderLimits } from './contract';

export const DEFAULT_LIMITS: RenderLimits = {
  maxInputBytes: 64 * 1024 * 1024,
  maxPdfBytes: 128 * 1024 * 1024,
  maxTotalImageBytes: 256 * 1024 * 1024,
  maxPages: 200,
  dpi: 120,
  thumbnailMaxEdge: 480,
};
export const DEFAULT_FEATURE_FLAGS: RenderFeatureFlags = {
  emitPdf: true,
  emitThumbnails: true,
  requireVisualReview: true,
};
export const DEFAULT_TIMEOUT_MS = 120_000;

export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  return {
    engine: { driver: 'libreoffice-poppler', sofficePath: 'soffice', pdftoppmPath: 'pdftoppm', ...overrides.engine },
    limits: { ...DEFAULT_LIMITS, ...overrides.limits },
    timeoutMs: overrides.timeoutMs ?? DEFAULT_TIMEOUT_MS,
    featureFlags: { ...DEFAULT_FEATURE_FLAGS, ...overrides.featureFlags },
  };
}
export function withCallOverrides(config: ModuleConfig, overrides?: ConfigOverrides): ModuleConfig {
  if (!overrides) return config;
  return {
    ...config,
    engine: { ...config.engine, ...overrides.engine },
    limits: { ...config.limits, ...overrides.limits },
    timeoutMs: overrides.timeoutMs ?? config.timeoutMs,
    featureFlags: { ...config.featureFlags, ...overrides.featureFlags },
  };
}

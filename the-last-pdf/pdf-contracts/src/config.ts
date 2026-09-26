import { configSchema } from './schema';
import { PdfModuleError } from './errors';
import type { ConfigOverrides, ModuleConfig } from './contract';

export const DEFAULT_CONFIG: Readonly<ModuleConfig> = Object.freeze({
  timeoutMs: 30_000,
  limits: Object.freeze({ maxInputBytes: 64 * 1024 * 1024, maxOutputJsonBytes: 16 * 1024 * 1024,
    maxPages: 200, maxNodes: 100_000, maxArtifacts: 500,
    maxOutputBytes: 128 * 1024 * 1024, maxTotalOutputBytes: 256 * 1024 * 1024, maxPagePixels: 25_000_000 }),
});
export function resolveConfig(overrides: ConfigOverrides = {}): ModuleConfig {
  const parsed = configSchema.safeParse(overrides);
  if (!parsed.success) throw new PdfModuleError('INVALID_INPUT', parsed.error.message);
  return Object.freeze({ timeoutMs: parsed.data.timeoutMs ?? DEFAULT_CONFIG.timeoutMs,
    limits: Object.freeze({ ...DEFAULT_CONFIG.limits, ...parsed.data.limits }) });
}
export function withCallOverrides(config: ModuleConfig, overrides: ConfigOverrides = {}): ModuleConfig {
  const parsed = configSchema.safeParse(overrides);
  if (!parsed.success) throw new PdfModuleError('INVALID_INPUT', parsed.error.message);
  if (overrides.timeoutMs !== undefined && overrides.timeoutMs > config.timeoutMs)
    throw new PdfModuleError('INVALID_INPUT', 'Request cannot relax the Profile timeout.');
  for (const [key, value] of Object.entries(overrides.limits ?? {})) {
    if (value! > config.limits[key as keyof ModuleConfig['limits']])
      throw new PdfModuleError('INVALID_INPUT', `Request cannot relax ${key}.`);
  }
  return resolveConfig({ timeoutMs: overrides.timeoutMs ?? config.timeoutMs, limits: { ...config.limits, ...overrides.limits } });
}

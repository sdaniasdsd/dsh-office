import { z } from 'zod';
import type { ConfigOverrides, ModuleConfig } from './contract';
import { failure } from './errors';
const positive = z.number().int().positive().max(512 * 1024 * 1024);
export const configSchema = z.object({
  engine: z.object({ driver: z.literal('docx4j-core-ts') }).strict().optional(),
  limits: z.object({
    maxInputBytes: positive.optional(), maxOutputBytes: positive.optional(), maxExpandedBytes: positive.optional(),
    maxEntries: positive.optional(), maxBlocks: positive.optional(), maxTableCells: positive.optional(), maxTextChars: positive.optional(),
  }).strict().optional(),
  timeoutMs: z.number().int().positive().max(600_000).optional(),
  featureFlags: z.object({ allowTemplateFill: z.boolean().optional(), allowImages: z.boolean().optional() }).strict().optional(),
}).strict();
export const DEFAULT_CONFIG: ModuleConfig = {
  engine: { driver: 'docx4j-core-ts' },
  limits: { maxInputBytes: 16 * 1024 * 1024, maxOutputBytes: 32 * 1024 * 1024, maxExpandedBytes: 64 * 1024 * 1024,
    maxEntries: 2048, maxBlocks: 2000, maxTableCells: 10_000, maxTextChars: 1_000_000 },
  timeoutMs: 30_000, featureFlags: { allowTemplateFill: true, allowImages: true },
};
export function resolveConfig(overrides: ConfigOverrides = {}, base = DEFAULT_CONFIG, tightenOnly = false): ModuleConfig {
  const parsed = configSchema.safeParse(overrides);
  if (!parsed.success) failure('INVALID_INPUT', 'Invalid module configuration.');
  const value = parsed.data;
  const result = { engine: { ...base.engine, ...value.engine }, limits: { ...base.limits, ...value.limits },
    timeoutMs: value.timeoutMs ?? base.timeoutMs, featureFlags: { ...base.featureFlags, ...value.featureFlags } };
  if (tightenOnly) {
    for (const key of Object.keys(base.limits) as (keyof ModuleConfig['limits'])[])
      if (result.limits[key] > base.limits[key]) failure('SAFETY_POLICY_DENIED', 'Per-call limits cannot exceed Profile limits.');
    if (result.timeoutMs > base.timeoutMs) failure('SAFETY_POLICY_DENIED', 'Per-call timeout cannot exceed Profile timeout.');
    for (const key of Object.keys(base.featureFlags) as (keyof ModuleConfig['featureFlags'])[])
      if (result.featureFlags[key] && !base.featureFlags[key]) failure('SAFETY_POLICY_DENIED', 'Per-call feature cannot override Profile policy.');
  }
  return result;
}

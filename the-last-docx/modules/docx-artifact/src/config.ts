import { z } from 'zod';
import type { ModuleConfig, ConfigOverrides } from './contract';
import { fail } from './errors';
const budget = z.number().int().positive().max(1024*1024*1024);
export const configSchema = z.object({
  engine:z.object({driver:z.literal('office-files-preview')}).strict().optional(),
  timeoutMs:z.number().int().positive().max(600000).optional(),
  limits:z.object({maxArtifactBytes:budget.optional(),maxManifestBytes:budget.optional(),maxTotalBytes:budget.optional(),maxReferences:budget.optional(),maxPages:budget.optional()}).strict().optional(),
  featureFlags:z.object({requirePreview:z.boolean().optional(),requireVisualReview:z.boolean().optional(),allowFailedReview:z.boolean().optional()}).strict().optional(),
}).strict();
export const DEFAULT_CONFIG:ModuleConfig = {engine:{driver:'office-files-preview'},timeoutMs:30000,
  limits:{maxArtifactBytes:64*1024*1024,maxManifestBytes:2*1024*1024,maxTotalBytes:512*1024*1024,maxReferences:1000,maxPages:300},
  featureFlags:{requirePreview:false,requireVisualReview:false,allowFailedReview:false}};
export function resolveConfig(value:ConfigOverrides={},base=DEFAULT_CONFIG,tighten=false):ModuleConfig {
  const parsed=configSchema.safeParse(value); if(!parsed.success) fail('INVALID_INPUT','Invalid delivery configuration.');
  const cfg={engine:{...base.engine,...parsed.data.engine},limits:{...base.limits,...parsed.data.limits},timeoutMs:parsed.data.timeoutMs??base.timeoutMs,featureFlags:{...base.featureFlags,...parsed.data.featureFlags}};
  if(tighten) {
    for(const key of Object.keys(base.limits) as (keyof ModuleConfig['limits'])[]) if(cfg.limits[key]>base.limits[key]) fail('SAFETY_POLICY_DENIED','Call cannot increase Profile budgets.');
    if(cfg.timeoutMs>base.timeoutMs || (!cfg.featureFlags.requirePreview&&base.featureFlags.requirePreview)||(!cfg.featureFlags.requireVisualReview&&base.featureFlags.requireVisualReview)||(cfg.featureFlags.allowFailedReview&&!base.featureFlags.allowFailedReview)) fail('SAFETY_POLICY_DENIED','Call cannot loosen Profile policy.');
  }
  return cfg;
}

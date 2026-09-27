import type { ConfigOverrides, ModuleConfig, RenderFeatureFlags, RenderLimits } from './contract';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { resolveRuntime } from '@dsh-office-profile/docx-runtime';

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
  const callerPath = fileURLToPath(import.meta.url);
  const runtime = resolveRuntime({
    runtimeRoot: overrides.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames,
    callerPath,
    bundledRuntimeRoot: join(dirname(callerPath), '..', 'runtime', 'win32-x64'),
    components: [
      { id: 'libreoffice', envVar: 'DOCX_SOFFICE', defaultEntry: 'libreoffice/program/soffice.com' },
      { id: 'poppler', envVar: 'DOCX_PDFTOPPM', defaultEntry: 'poppler/poppler-26.09.0/Library/bin/pdftoppm.exe' },
    ],
  });
  const engine = { driver: 'libreoffice-poppler' as const, sofficePath: 'soffice', pdftoppmPath: 'pdftoppm', ...overrides.engine };
  if (overrides.engine?.sofficePath === undefined && runtime.components.libreoffice?.path !== undefined) {
    engine.sofficePath = runtime.components.libreoffice.path;
  }
  if (overrides.engine?.pdftoppmPath === undefined && runtime.components.poppler?.path !== undefined) {
    engine.pdftoppmPath = runtime.components.poppler.path;
  }
  return {
    engine,
    runtimeRoot: overrides.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames,
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
    runtimeRoot: overrides.runtimeRoot ?? config.runtimeRoot,
    runtimePackageNames: overrides.runtimePackageNames ?? config.runtimePackageNames,
    limits: { ...config.limits, ...overrides.limits },
    timeoutMs: overrides.timeoutMs ?? config.timeoutMs,
    featureFlags: { ...config.featureFlags, ...overrides.featureFlags },
  };
}

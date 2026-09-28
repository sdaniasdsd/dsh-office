import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { ExecuteResult as DocxParseExecuteResult } from '@dsh-office-profile/docx-parse';

export const DOCX_RENDER_MODULE_ID = 'docx-render' as const;
export const DOCX_RENDER_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export type DocxRenderCapability = (typeof DOCX_RENDER_CAPABILITIES)[number];

export const DOCX_RENDER_ERROR_CODES = {
  INVALID_INPUT: 'INVALID_INPUT',
  ARTIFACT_NOT_FOUND: 'ARTIFACT_NOT_FOUND',
  UNSUPPORTED_ARTIFACT_URI: 'UNSUPPORTED_ARTIFACT_URI',
  FORMAT_MISMATCH: 'FORMAT_MISMATCH',
  SAFETY_POLICY_DENIED: 'SAFETY_POLICY_DENIED',
  ARTIFACT_STORE_UNAVAILABLE: 'ARTIFACT_STORE_UNAVAILABLE',
  ENGINE_UNAVAILABLE: 'ENGINE_UNAVAILABLE',
  ENGINE_FAILED: 'ENGINE_FAILED',
  ENGINE_TIMEOUT: 'ENGINE_TIMEOUT',
  ENGINE_PROTOCOL_ERROR: 'ENGINE_PROTOCOL_ERROR',
  LIMIT_EXCEEDED: 'LIMIT_EXCEEDED',
  SOURCE_MAP_REQUIRED: 'SOURCE_MAP_REQUIRED',
  SOURCE_MAP_MISMATCH: 'SOURCE_MAP_MISMATCH',
  STALE_RENDER_RESULT: 'STALE_RENDER_RESULT',
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;
export type DocxRenderErrorCode = keyof typeof DOCX_RENDER_ERROR_CODES;

export interface RenderLimits {
  maxInputBytes: number;
  maxPdfBytes: number;
  maxTotalImageBytes: number;
  maxPages: number;
  dpi: number;
  thumbnailMaxEdge: number;
}
export interface RenderFeatureFlags {
  emitPdf: boolean;
  emitThumbnails: boolean;
  requireVisualReview: boolean;
}
export interface RenderEngineConfig {
  driver: 'libreoffice-poppler';
  sofficePath: string;
  pdftoppmPath: string;
  tempRoot?: string;
}
export interface ModuleConfig {
  engine: RenderEngineConfig;
  /** Optional DSH runtime root; resolved before falling back to PATH commands. */
  runtimeRoot?: string;
  /** Additional sibling runtime packages for a host-specific deployment. */
  runtimePackageNames?: readonly string[];
  limits: RenderLimits;
  timeoutMs: number;
  featureFlags: RenderFeatureFlags;
}
export interface ConfigOverrides {
  engine?: Partial<RenderEngineConfig>;
  runtimeRoot?: string;
  runtimePackageNames?: readonly string[];
  limits?: Partial<RenderLimits>;
  timeoutMs?: number;
  featureFlags?: Partial<RenderFeatureFlags>;
}
export interface DocxRenderOptions extends ConfigOverrides {}

interface ModuleInputBase {
  artifactRef: ArtifactRef;
  requestId: string;
  options?: DocxRenderOptions;
  policy?: SafetyPolicy;
}
export interface InspectInput extends ModuleInputBase { operation: 'inspect' }
export interface ExecuteInput extends ModuleInputBase { operation: 'execute' }

/** Findings are observations from the agent/human reviewing rendered page images, not parser errors. */
export interface VisualFinding {
  id: string;
  kind: 'truncation' | 'overlap' | 'pagination' | 'missing-content' | 'other';
  severity: 'error' | 'warn' | 'info';
  message: string;
  page: number;
  confidence?: number;
  /** Optional references into docx-parse's dual IR; docx-render does not copy or reinterpret its schema. */
  sourceSemanticIds?: string[];
  sourcePointers?: string[];
}
export interface VerifyInput extends ModuleInputBase {
  operation: 'verify';
  /** Supply the render result from execute to avoid rendering the same source twice. */
  renderResult?: RenderResult;
  /** Public execute result from docx-parse; required only when findings carry source references. */
  parseResult?: DocxParseExecuteResult;
  /** Pass findings after reviewing the returned PNGs; absent findings are not treated as a visual pass. */
  visualFindings?: VisualFinding[];
}

export interface VerificationCheck {
  id: string;
  status: 'pass' | 'fail' | 'skip';
  severity: 'error' | 'warn' | 'info';
  message: string;
  details?: Record<string, string | number | boolean>;
}
export interface VerificationReport {
  ok: boolean;
  partial: boolean;
  checks: VerificationCheck[];
  summary: { total: number; passed: number; failed: number; skipped: number };
}
export interface RenderPage {
  pageNumber: number;
  image: ArtifactRef;
  thumbnail?: ArtifactRef;
}
export interface RenderResult {
  engine: string;
  sourceArtifactId: string;
  sourceSha256?: string;
  pageCount: number;
  pdf?: ArtifactRef;
  pages: RenderPage[];
  visualReview: 'pending' | 'provided';
  findings: VisualFinding[];
}
export interface ArtifactStore {
  read(ref: ArtifactRef, maxBytes: number): Promise<Uint8Array>;
  write(input: { source: ArtifactRef; requestId: string; bytes: Uint8Array; suggestedName: string }): Promise<ArtifactRef>;
}
export interface SafetyGuard {
  assertAllowed(input: { artifactRef: ArtifactRef; policy: SafetyPolicy; operation: DocxRenderCapability }): Promise<void>;
}
export interface TelemetrySink {
  record(event: { name: string; timestampMs: number; durationMs: number; attributes: Record<string, string | number | boolean> }): void;
}
export interface ModuleDependencyRef { name: string; kind: 'module' | 'runtime'; optional?: boolean; version?: string }
export interface DocxRenderModuleDefinition {
  id: typeof DOCX_RENDER_MODULE_ID;
  version: string;
  profileGroup: 'DOCX';
  summary: string;
  capabilities: typeof DOCX_RENDER_CAPABILITIES;
  dependencies: readonly ModuleDependencyRef[];
  configSchema: Record<string, unknown>;
}
export interface DocxRenderHandlers {
  inspect(input: InspectInput): Promise<ModuleOutput<FormatProfile>>;
  execute(input: ExecuteInput): Promise<ModuleOutput<RenderResult>>;
  verify(input: VerifyInput): Promise<ModuleOutput<VerificationReport> & { verification: VerificationReport }>;
}
export interface ModuleOutput<T> {
  moduleId: typeof DOCX_RENDER_MODULE_ID;
  requestId: string;
  operation: DocxRenderCapability;
  result: T;
  artifacts: ArtifactRef[];
  warnings: Warning[];
  verification?: VerificationReport;
}
export interface DocxRenderModule { definition: DocxRenderModuleDefinition; handlers: DocxRenderHandlers; dispose(): Promise<void> }
export interface ModuleCreateOptions {
  config?: ConfigOverrides;
  artifactStore: ArtifactStore;
  safetyGuard?: SafetyGuard;
  telemetry?: TelemetrySink;
  engine?: RenderEngine;
}
export interface EnginePage { pageNumber: number; image: Uint8Array; thumbnail?: Uint8Array }
export interface EngineRenderResult { pdf: Uint8Array; pages: EnginePage[] }
export interface RenderEngine {
  readonly name: string;
  inspect(bytes: Uint8Array, artifactRef: ArtifactRef): Promise<{ extension: string | null; mediaType: string | null }>;
  render(bytes: Uint8Array, config: ModuleConfig, artifactRef: ArtifactRef): Promise<EngineRenderResult>;
  dispose(): Promise<void>;
}

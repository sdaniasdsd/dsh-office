import type { ArtifactRef, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { StyleInventory, StylePlan, StyleSpec } from './domain/docx-styles';

export const DOCX_STYLES_MODULE_ID = 'docx-styles' as const;
export const DOCX_STYLES_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export type DocxStylesCapability = (typeof DOCX_STYLES_CAPABILITIES)[number];

export const DOCX_STYLES_ERROR_CODES = {
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
  UNSUPPORTED_OPERATION: 'UNSUPPORTED_OPERATION',
  RELATIONSHIP_INTEGRITY_FAILED: 'RELATIONSHIP_INTEGRITY_FAILED',
  LIMIT_EXCEEDED: 'LIMIT_EXCEEDED',
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;
export type DocxStylesErrorCode = keyof typeof DOCX_STYLES_ERROR_CODES;

export interface StyleLimits {
  maxInputBytes: number;
  maxOutputBytes: number;
  maxStyles: number;
}

export interface StyleFeatureFlags {
  allowStyleDefinitions: boolean;
  allowRedefineExisting: boolean;
}

export interface ModuleConfig {
  engine: { driver: 'docx4j-core-ts' };
  limits: StyleLimits;
  timeoutMs: number;
  featureFlags: StyleFeatureFlags;
}

export interface ConfigOverrides {
  limits?: Partial<StyleLimits>;
  timeoutMs?: number;
  featureFlags?: Partial<StyleFeatureFlags>;
}

export interface DocxStylesOptions extends ConfigOverrides {}

interface ModuleInputBase {
  artifactRef: ArtifactRef;
  requestId: string;
  options?: DocxStylesOptions;
}

export interface InspectInput extends ModuleInputBase {
  operation: 'inspect';
  policy?: SafetyPolicy;
}

export interface ExecuteInput extends ModuleInputBase {
  operation: 'execute';
  plan: StylePlan;
}

export interface VerifyInput extends ModuleInputBase {
  operation: 'verify';
  policy?: SafetyPolicy;
  /** The definitions the caller expects to find; omitted means "audit only". */
  expectations?: { styles?: readonly StyleSpec[] };
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

export interface StyleResult {
  artifact: ArtifactRef;
  defined: StyleDefinitionResult[];
  inventory: StyleInventory;
  verification: VerificationReport;
}

export interface StyleDefinitionResult {
  styleId: string;
  name: string;
  type: 'paragraph' | 'character' | 'table';
  action: 'created' | 'replaced' | 'unchanged';
  properties: string[];
}

export interface ModuleOutput<T> {
  moduleId: typeof DOCX_STYLES_MODULE_ID;
  requestId: string;
  operation: DocxStylesCapability;
  result: T;
  artifacts: ArtifactRef[];
  warnings: Warning[];
  verification?: VerificationReport;
}

export type InspectOutput = ModuleOutput<StyleInventory>;
export type ExecuteOutput = ModuleOutput<StyleResult>;
export type VerifyOutput = ModuleOutput<VerificationReport>;

export interface ArtifactStore {
  read(ref: ArtifactRef, maxBytes: number): Promise<Uint8Array>;
  write(input: {
    source: ArtifactRef;
    requestId: string;
    bytes: Uint8Array;
    suggestedName: string;
  }): Promise<ArtifactRef>;
}

export interface SafetyGuard {
  assertAllowed(input: { artifactRef: ArtifactRef; policy: SafetyPolicy; operation: DocxStylesCapability }): Promise<void>;
}

export interface TelemetrySink {
  record(event: {
    name: string;
    timestampMs: number;
    durationMs: number;
    attributes: Record<string, string | number | boolean>;
  }): void;
}

export interface ModuleDependencyRef {
  name: string;
  kind: 'module' | 'npm';
  optional?: boolean;
  version?: string;
}

export interface DocxStylesModuleDefinition {
  id: typeof DOCX_STYLES_MODULE_ID;
  version: string;
  profileGroup: 'DOCX';
  summary: string;
  capabilities: typeof DOCX_STYLES_CAPABILITIES;
  dependencies: readonly ModuleDependencyRef[];
  configSchema: Record<string, unknown>;
}

export interface DocxStylesHandlers {
  inspect(input: InspectInput): Promise<InspectOutput>;
  execute(input: ExecuteInput): Promise<ExecuteOutput>;
  verify(input: VerifyInput): Promise<VerifyOutput>;
}

export interface DocxStylesModule {
  definition: DocxStylesModuleDefinition;
  handlers: DocxStylesHandlers;
  dispose(): Promise<void>;
}

export interface ModuleCreateOptions {
  config?: ConfigOverrides;
  artifactStore: ArtifactStore;
  safetyGuard?: SafetyGuard;
  telemetry?: TelemetrySink;
}

export type { StyleInventory, StylePlan, StyleSpec };

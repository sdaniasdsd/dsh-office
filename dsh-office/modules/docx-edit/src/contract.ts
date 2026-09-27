import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { SemanticId } from '@dsh-office-profile/docx-parse';
import type { AppliedEdit, DocxEditOperation, EditPlan } from './domain/docx-edit';

export const DOCX_EDIT_MODULE_ID = 'docx-edit' as const;
export const DOCX_EDIT_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export type DocxEditCapability = (typeof DOCX_EDIT_CAPABILITIES)[number];

export const DOCX_EDIT_ERROR_CODES = {
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
  TARGET_NOT_FOUND: 'TARGET_NOT_FOUND',
  TARGET_AMBIGUOUS: 'TARGET_AMBIGUOUS',
  STALE_TARGET: 'STALE_TARGET',
  TEXT_NOT_FOUND: 'TEXT_NOT_FOUND',
  TEXT_AMBIGUOUS: 'TEXT_AMBIGUOUS',
  UNSUPPORTED_OPERATION: 'UNSUPPORTED_OPERATION',
  RELATIONSHIP_INTEGRITY_FAILED: 'RELATIONSHIP_INTEGRITY_FAILED',
  LIMIT_EXCEEDED: 'LIMIT_EXCEEDED',
  VERIFICATION_FAILED: 'VERIFICATION_FAILED',
} as const;
export type DocxEditErrorCode = keyof typeof DOCX_EDIT_ERROR_CODES;

export interface EditLimits {
  maxInputBytes: number;
  maxOutputBytes: number;
  maxOperations: number;
}

export interface EditFeatureFlags {
  allowTextReplacement: boolean;
  allowComments: boolean;
  allowTrackedChanges: boolean;
  allowAcceptReject: boolean;
  requireExpectedDigest: boolean;
  /** In-place paragraph formatting. Additive: turning it off leaves every other edit kind untouched. */
  allowFormatting: boolean;
  /**
   * Table instance geometry (`w:tblW`, grid widths, cell margins, borders).
   * Separate from `allowFormatting` because it rewrites the table element rather
   * than one paragraph's properties, and separate from the style module because
   * geometry is content, not a reusable definition.
   */
  allowTableFormatting: boolean;
  /**
   * Adding paragraphs. Gated separately because it changes the document's block
   * structure rather than the properties of blocks that already exist.
   */
  allowBlockEdits: boolean;
}

export interface ModuleConfig {
  engine: { driver: 'docx4j-core-ts' };
  limits: EditLimits;
  timeoutMs: number;
  featureFlags: EditFeatureFlags;
}

export interface ConfigOverrides {
  limits?: Partial<EditLimits>;
  timeoutMs?: number;
  featureFlags?: Partial<EditFeatureFlags>;
}

export interface DocxEditOptions extends ConfigOverrides {}

interface ModuleInputBase {
  artifactRef: ArtifactRef;
  requestId: string;
  options?: DocxEditOptions;
}

export interface InspectInput extends ModuleInputBase {
  operation: 'inspect';
  policy?: SafetyPolicy;
}

export interface ExecuteInput extends ModuleInputBase {
  operation: 'execute';
  plan: EditPlan;
}

export interface VerifyInput extends ModuleInputBase {
  operation: 'verify';
  policy?: SafetyPolicy;
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

export interface EditResult {
  artifact: ArtifactRef;
  edits: AppliedEdit[];
  verification: VerificationReport;
}

export interface ModuleOutput<T> {
  moduleId: typeof DOCX_EDIT_MODULE_ID;
  requestId: string;
  operation: DocxEditCapability;
  result: T;
  artifacts: ArtifactRef[];
  warnings: Warning[];
  verification?: VerificationReport;
}

export type InspectOutput = ModuleOutput<FormatProfile>;
export type ExecuteOutput = ModuleOutput<EditResult>;
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
  assertAllowed(input: { artifactRef: ArtifactRef; policy: SafetyPolicy; operation: DocxEditCapability }): Promise<void>;
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

export interface DocxEditModuleDefinition {
  id: typeof DOCX_EDIT_MODULE_ID;
  version: string;
  profileGroup: 'DOCX';
  summary: string;
  capabilities: typeof DOCX_EDIT_CAPABILITIES;
  dependencies: readonly ModuleDependencyRef[];
  configSchema: Record<string, unknown>;
}

export interface DocxEditHandlers {
  inspect(input: InspectInput): Promise<InspectOutput>;
  execute(input: ExecuteInput): Promise<ExecuteOutput>;
  verify(input: VerifyInput): Promise<VerifyOutput>;
}

export interface DocxEditModule {
  definition: DocxEditModuleDefinition;
  handlers: DocxEditHandlers;
  dispose(): Promise<void>;
}

export interface ModuleCreateOptions {
  config?: ConfigOverrides;
  artifactStore: ArtifactStore;
  safetyGuard?: SafetyGuard;
  telemetry?: TelemetrySink;
}

export type { SemanticId, DocxEditOperation };

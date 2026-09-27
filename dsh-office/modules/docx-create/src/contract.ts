import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { VerificationReportLike } from 'office-test-kit';
import type { CreatePlan, DesignDecision, DocumentSpec } from './domain/docx-create';

export const DOCX_CREATE_MODULE_ID = 'docx-create' as const;
export interface ModuleConfig {
  engine: { driver: 'docx4j-core-ts' };
  limits: { maxInputBytes: number; maxOutputBytes: number; maxExpandedBytes: number; maxEntries: number; maxBlocks: number; maxTableCells: number; maxTextChars: number };
  timeoutMs: number;
  featureFlags: { allowTemplateFill: boolean; allowImages: boolean };
}
export interface ConfigOverrides {
  engine?: ModuleConfig['engine']; limits?: Partial<ModuleConfig['limits']>;
  timeoutMs?: number; featureFlags?: Partial<ModuleConfig['featureFlags']>;
}
interface BaseInput { requestId: string; options?: ConfigOverrides; policy?: SafetyPolicy }
export interface InspectInput extends BaseInput { operation: 'inspect'; artifactRef: ArtifactRef }
export interface VerifyInput extends BaseInput { operation: 'verify'; artifactRef: ArtifactRef }
export interface ExecuteInput extends BaseInput {
  operation: 'execute';
  plan: { kind: 'create'; document: DocumentSpec } | Extract<CreatePlan, { kind: 'fillTemplate' }>;
}
export interface ArtifactStore {
  read(ref: ArtifactRef, maxBytes: number): Promise<Uint8Array>;
  /** Must create a NEW immutable artifact, never overwrite any input. No source exists for blank creation. */
  write(input: { requestId: string; sources: ArtifactRef[]; bytes: Uint8Array; suggestedName: string }): Promise<ArtifactRef>;
}
export interface VerificationCheck { id: string; status: 'pass' | 'fail' | 'skip'; severity: 'error' | 'warn' | 'info'; message: string }
export interface VerificationReport extends VerificationReportLike { checks: VerificationCheck[] }
export interface CreateResult {
  artifactRef: ArtifactRef; engine: string; mode: CreatePlan['kind'];
  paragraphIds: string[]; filledKeys: string[];
  fields: 'pending-update' | 'none'; visualReview: 'pending';
  /**
   * Present for `create`. Absent for `fillTemplate`, where the template already
   * owns its own styling and there is no decoration decision to make.
   */
  design?: DesignDecision;
}
export interface ModuleOutput<T> {
  moduleId: 'docx-create'; requestId: string; operation: 'inspect' | 'execute' | 'verify';
  result: T; artifacts: ArtifactRef[]; warnings: Warning[]; verification?: VerificationReport;
}
/** Host-side engine port only. Bytes never appear in serialized handler results. */
export interface EngineRequest {
  plan: CreatePlan; template?: Uint8Array; images: Record<string, Uint8Array>; config: ModuleConfig;
}
export interface EngineResult { bytes: Uint8Array; filledKeys: string[]; design?: DesignDecision }
export interface CreateEngine {
  readonly name: string;
  execute(input: EngineRequest): Promise<EngineResult>;
  dispose(): Promise<void>;
}
export interface TelemetrySink { record(event: { moduleId: 'docx-create'; operation: string; elapsedMs: number; ok: boolean; errorCode?: string }): void | Promise<void> }
export interface ModuleCreateOptions {
  artifactStore: ArtifactStore; config?: ConfigOverrides; engine?: CreateEngine;
  safetyGuard?: { assertAllowed(input: { artifactRef?: ArtifactRef; policy?: SafetyPolicy; operation: string }): Promise<void> };
  telemetry?: TelemetrySink;
}
export interface DocxCreateModule {
  definition: { id: string; version: string; capabilities: readonly string[]; profileGroup: string; configSchema: Record<string, unknown>; dependencies: readonly { name: string; kind: string; optional?: boolean }[] };
  handlers: {
    inspect(input: InspectInput): Promise<ModuleOutput<FormatProfile>>;
    execute(input: ExecuteInput): Promise<ModuleOutput<CreateResult>>;
    verify(input: VerifyInput): Promise<ModuleOutput<VerificationReport>>;
  };
  dispose(): Promise<void>;
}

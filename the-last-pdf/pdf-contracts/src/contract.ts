import type { ArtifactRef, Warning } from 'office-core';
import type { VerificationReportLike } from 'office-test-kit';
import type { z } from 'zod';
import type { configSchema, requestSchema, profileSchema, policySchema, checkSchema, dualIRSchema, limitsSchema } from './schema';
export type { ArtifactRef, Warning };
export type PdfFormatProfile = z.infer<typeof profileSchema>;
export type PdfSafetyPolicy = z.infer<typeof policySchema>;
export type PdfDualIR = z.infer<typeof dualIRSchema>;
export type ConfigOverrides = z.infer<typeof configSchema>;
export type ModuleRequest = z.infer<typeof requestSchema>;
export type VerificationCheck = z.infer<typeof checkSchema>;
export interface VerificationReport extends VerificationReportLike { checks: VerificationCheck[] }
export interface ModuleConfig { timeoutMs: number; limits: z.infer<typeof limitsSchema> }
export type Operation = ModuleRequest['operation'];
export type ModuleId = 'pdf-inspect' | 'pdf-easy-parse' | 'pdf-parse' | 'pdf-complex-parse'
  | 'pdf-create' | 'pdf-edit' | 'pdf-render' | 'pdf-artifact';
export interface ModuleDefinition {
  id: ModuleId; name: string; version: string; profileGroup: 'PDF'; entry: string;
  summary: string; capabilities: readonly ['inspect', 'execute', 'verify'];
  implementation: 'framework-only' | 'implemented'; engineSelection: 'pending-user-decision' | 'user-approved';
  decisionIds: readonly string[];
}
export interface ArtifactReader {
  /** Host enforces the budget before allocating the entire file. */
  read(ref: ArtifactRef, maxBytes: number, signal: AbortSignal): Promise<Uint8Array>;
}
/** Reserved persistence port for chosen write adapters. Host owns immutable writes/CAS.
 * Do not bind this port to an engine until that engine's write semantics are implemented. */
export interface ArtifactWriter {
  write(input: { sources: ArtifactRef[]; requestId: string; bytes: Uint8Array;
    suggestedName: string; signal: AbortSignal }): Promise<ArtifactRef>;
  commitManifest(input: { documentId: string; revision: number; parentSha256?: string;
    requestId: string; bytes: Uint8Array; signal: AbortSignal }): Promise<ArtifactRef>;
}
export interface EngineContext {
  moduleId: ModuleId; request: ModuleRequest; config: ModuleConfig;
  source?: { ref: ArtifactRef; bytes: Uint8Array; identity: { id: string; sha256: string } };
  signal: AbortSignal;
}
export interface PdfEngine {
  name: string;
  /** Return { result, artifacts: [], warnings: [] }. Only JSON across the boundary.
   * Child-process engines must kill/reap their process on abort, cap stdout/stderr,
   * and report actual feature coverage. No engine is installed by this framework. */
  invoke(context: EngineContext): Promise<unknown>;
  dispose?(): Promise<void>;
}
export interface ModuleOptions {
  engine?: PdfEngine;
  files?: ArtifactReader;
  writer?: ArtifactWriter;
  backend?: {
    qpdfPath?: string;
    requireQpdf?: boolean;
    fontPath?: string;
    pythonPath?: string;
    modelPath?: string;
    doclingScript?: string;
  };
  config?: ConfigOverrides;
  safetyGuard?: { assertAllowed(context: EngineContext & { policy: PdfSafetyPolicy }): Promise<void> };
  telemetry?: { record(event: { moduleId: ModuleId; operation: Operation; ok: boolean;
    elapsedMs: number; errorCode?: string }): void | Promise<void> };
}
export interface ModuleOutput<T> {
  moduleId: ModuleId; requestId: string; operation: Operation;
  result: T; artifacts: ArtifactRef[]; warnings: Warning[];
  verification?: VerificationReport;
}

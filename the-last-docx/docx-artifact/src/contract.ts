import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import type { VerificationReportLike } from 'office-test-kit';
import type { RenderResult } from '@dsh-office-profile/docx-render';
export type { ArtifactRef, RenderResult };
export const DOCX_ARTIFACT_MODULE_ID = 'docx-artifact' as const;
export interface ModuleConfig {
  engine: { driver: 'office-files-preview' }; timeoutMs: number;
  limits: { maxArtifactBytes: number; maxManifestBytes: number; maxTotalBytes: number; maxReferences: number; maxPages: number };
  featureFlags: { requirePreview: boolean; requireVisualReview: boolean; allowFailedReview: boolean };
}
export interface ConfigOverrides { timeoutMs?: number; limits?: Partial<ModuleConfig['limits']>; featureFlags?: Partial<ModuleConfig['featureFlags']>; engine?: ModuleConfig['engine'] }
export type BridgeKind = 'inspection' | 'dual-ir' | 'easy-ir' | 'complex-ir' | 'node-bridge' | 'verification';
export interface BridgeReference { kind: BridgeKind; artifactRef: ArtifactRef; sourceArtifactId: string; sourceSha256: string }
export interface DeliveryPlan {
  documentId: string; revision: number; parentManifest?: ArtifactRef;
  preview?: RenderResult; bridges?: BridgeReference[];
}
export interface DeliveryManifest {
  schema: 'docx-delivery/v1'; documentId: string; revision: number;
  document: ArtifactRef; parentManifest?: ArtifactRef; preview?: RenderResult;
  bridges: BridgeReference[];
  review: { state: 'pending' | 'reviewed' | 'failed'; basis: 'none' | 'caller-provided-findings' };
}
export interface VerificationCheck { id: string; status: 'pass'|'fail'|'skip'; severity: 'error'|'warn'|'info'; message: string }
export interface VerificationReport extends VerificationReportLike { checks: VerificationCheck[] }
interface Base { artifactRef: ArtifactRef; requestId: string; options?: ConfigOverrides; policy?: SafetyPolicy }
export interface InspectInput extends Base { operation: 'inspect' }
export interface ExecuteInput extends Base { operation: 'execute'; delivery: DeliveryPlan }
/** verify consumes the manifest ArtifactRef, not the DOCX reference. */
export interface VerifyInput extends Base { operation: 'verify' }
export interface ModuleOutput<T> { moduleId:'docx-artifact'; operation:'inspect'|'execute'|'verify'; requestId:string; result:T; artifacts:ArtifactRef[]; warnings:Warning[]; verification?:VerificationReport }
export interface DeliveryResult { manifestRef: ArtifactRef; documentRef: ArtifactRef; manifest: DeliveryManifest }
/** Atomic commit/CAS belongs to office-files. Same document/revision + same bytes is idempotent;
 * different bytes conflict. No destructive replace, no last-write-wins version counter. */
export interface ArtifactFiles {
  read(ref: ArtifactRef, maxBytes: number, signal?: AbortSignal): Promise<Uint8Array>;
  commitManifest(input: { documentId:string; revision:number; parentSha256?:string; requestId:string; bytes:Uint8Array; signal:AbortSignal }): Promise<ArtifactRef>;
}
export interface ArtifactEngine {
  inspect(ref: ArtifactRef, signal: AbortSignal): Promise<FormatProfile>;
  read(ref: ArtifactRef, maxBytes: number, signal: AbortSignal): Promise<Uint8Array>;
  commit(input: Parameters<ArtifactFiles['commitManifest']>[0]): Promise<ArtifactRef>;
  dispose(): Promise<void>;
}
export interface ModuleCreateOptions {
  config?: ConfigOverrides; engine?: ArtifactEngine;
  files?: ArtifactFiles;
  inspector?: { inspect(ref:ArtifactRef, signal:AbortSignal):Promise<FormatProfile> };
  safetyGuard?: { assertAllowed(input:{artifactRef:ArtifactRef;operation:string;policy?:SafetyPolicy}):Promise<void> };
  telemetry?: { record(event:{moduleId:'docx-artifact';operation:string;ok:boolean;elapsedMs:number;errorCode?:string}):void|Promise<void> };
}
export interface DocxArtifactModule {
  definition: { id:string; version:string; profileGroup:string; capabilities:readonly string[]; configSchema:Record<string,unknown>; dependencies:readonly {name:string;kind:string;optional?:boolean}[] };
  handlers: {
    inspect(input:InspectInput):Promise<ModuleOutput<FormatProfile>>;
    execute(input:ExecuteInput):Promise<ModuleOutput<DeliveryResult>>;
    verify(input:VerifyInput):Promise<ModuleOutput<VerificationReport>>;
  };
  dispose():Promise<void>;
}

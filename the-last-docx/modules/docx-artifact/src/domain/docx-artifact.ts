import { z } from 'zod';
import type { ArtifactRef, DeliveryManifest, RenderResult } from '../contract';
import { configSchema } from '../config';
import { fail } from '../errors';
export const refSchema:z.ZodType<ArtifactRef>=z.object({id:z.string().min(1),uri:z.string().min(1),mediaType:z.string().optional(),sizeBytes:z.number().int().nonnegative().optional(),sha256:z.string().regex(/^[a-f0-9]{64}$/).optional(),label:z.string().optional(),tags:z.record(z.string()).optional()}).strict();
const digest=z.string().regex(/^[a-f0-9]{64}$/);
export const previewSchema:z.ZodType<RenderResult>=z.object({
  engine:z.string().min(1),sourceArtifactId:z.string().min(1),sourceSha256:digest.optional(),pageCount:z.number().int().positive(),
  pdf:refSchema.optional(),pages:z.array(z.object({pageNumber:z.number().int().positive(),image:refSchema,thumbnail:refSchema.optional()}).strict()),
  visualReview:z.enum(['pending','provided']),findings:z.array(z.object({id:z.string().min(1),kind:z.enum(['truncation','overlap','pagination','missing-content','other']),severity:z.enum(['error','warn','info']),message:z.string(),page:z.number().int().positive(),confidence:z.number().min(0).max(1).optional(),sourceSemanticIds:z.array(z.string()).optional(),sourcePointers:z.array(z.string()).optional()}).strict()),
}).strict();
export const bridgeSchema=z.object({kind:z.enum(['inspection','dual-ir','easy-ir','complex-ir','node-bridge','verification']),artifactRef:refSchema,sourceArtifactId:z.string().min(1),sourceSha256:digest}).strict();
const documentId=z.string().regex(/^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/);
export const planSchema=z.object({documentId,revision:z.number().int().positive(),parentManifest:refSchema.optional(),preview:previewSchema.optional(),bridges:z.array(bridgeSchema).optional()}).strict();
export const manifestSchema:z.ZodType<DeliveryManifest>=z.object({schema:z.literal('docx-delivery/v1'),documentId,revision:z.number().int().positive(),document:refSchema,parentManifest:refSchema.optional(),preview:previewSchema.optional(),bridges:z.array(bridgeSchema),review:z.object({state:z.enum(['pending','reviewed','failed']),basis:z.enum(['none','caller-provided-findings'])}).strict()}).strict();
const base={artifactRef:refSchema,requestId:z.string().trim().min(1).max(200),options:configSchema.optional(),policy:z.object({id:z.string().min(1)}).passthrough().optional()};
export const inspectSchema=z.object({...base,operation:z.literal('inspect')}).strict();
export const executeSchema=z.object({...base,operation:z.literal('execute'),delivery:planSchema}).strict();
export const verifySchema=z.object({...base,operation:z.literal('verify')}).strict();
export function jsonValue(value:unknown,seen=new Set<object>(),depth=0):void {
  if(depth>80) fail('INVALID_INPUT','JSON nesting exceeds limit.');
  if(value===null||typeof value==='string'||typeof value==='boolean'||(typeof value==='number'&&Number.isFinite(value))) return;
  if(typeof value!=='object'||seen.has(value)||(!Array.isArray(value)&&Object.getPrototypeOf(value)!==Object.prototype&&Object.getPrototypeOf(value)!==null)) fail('INVALID_INPUT','Value is not lossless JSON.');
  seen.add(value); for(const v of Object.values(value)) jsonValue(v,seen,depth+1); seen.delete(value);
}
export function parse<S extends z.ZodTypeAny>(schema:S,value:unknown):z.output<S> { jsonValue(value); const p=schema.safeParse(value);if(!p.success) fail('INVALID_INPUT',p.error.issues.map(x=>`${x.path.join('.')}: ${x.message}`).join('; '));return p.data; }
/** Deterministic JSON for idempotent immutable manifests (object keys sorted, array order retained). */
export function canonicalJson(value:unknown):string {
  jsonValue(value);
  const sort=(v:unknown):unknown=>Array.isArray(v)?v.map(sort):v!==null&&typeof v==='object'?Object.fromEntries(Object.keys(v).sort().map(k=>[k,sort((v as Record<string,unknown>)[k])])):v;
  return JSON.stringify(sort(value));
}

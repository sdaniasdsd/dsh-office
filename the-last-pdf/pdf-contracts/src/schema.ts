import { z } from 'zod';

const positive = z.number().int().positive().max(Number.MAX_SAFE_INTEGER);
const nonnegative = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER);
export const digestSchema = z.string().regex(/^[a-f0-9]{64}$/);
export const artifactSchema = z.object({
  id: z.string().min(1), uri: z.string().min(1), mediaType: z.string().optional(),
  sizeBytes: nonnegative.optional(), sha256: digestSchema.optional(),
  label: z.string().optional(), tags: z.record(z.string()).optional(),
}).strict();
export const sourceSchema = z.object({ id: z.string().min(1), sha256: digestSchema }).strict();
export const warningSchema = z.object({
  code: z.string().min(1), message: z.string().min(1),
  severity: z.enum(['info', 'warn', 'error']), path: z.string().optional(),
  details: z.record(z.unknown()).optional(),
}).strict();
export const checkSchema = z.object({
  id: z.string().min(1), status: z.enum(['pass', 'fail', 'skip']),
  severity: z.enum(['info', 'warn', 'error']), message: z.string().min(1),
}).strict();
export const reportSchema = z.object({
  ok: z.boolean(), partial: z.boolean(), checks: z.array(checkSchema),
  summary: z.object({ total: nonnegative, passed: nonnegative, failed: nonnegative, skipped: nonnegative }).strict(),
}).strict();
/** Unknown must stay null. A missing detector cannot claim absence. */
export const profileSchema = z.object({
  format: z.literal('pdf'), mediaType: z.literal('application/pdf'),
  container: z.literal('pdf'), version: z.string().nullable(),
  pageCount: positive.nullable(), encrypted: z.boolean().nullable(),
  features: z.object({
    javascript: z.boolean().nullable(), embeddedFiles: z.boolean().nullable(),
    externalLinks: z.boolean().nullable(), acroForm: z.boolean().nullable(),
    xfa: z.boolean().nullable(), signatures: z.boolean().nullable(),
    tagged: z.boolean().nullable(),
  }).strict(),
  coverage: z.enum(['observed', 'partial', 'unavailable']),
}).strict();
export const policySchema = z.object({
  id: z.string().min(1), allowEncrypted: z.boolean().optional(),
  allowJavaScript: z.boolean().optional(), allowEmbeddedFiles: z.boolean().optional(),
  allowExternalLinks: z.boolean().optional(), maxPages: positive.optional(),
}).strict();
export const limitsSchema = z.object({
  maxInputBytes: positive, maxOutputJsonBytes: positive, maxPages: positive,
  maxNodes: positive, maxArtifacts: positive,
  maxOutputBytes: positive, maxTotalOutputBytes: positive, maxPagePixels: positive,
}).strict();
export const configSchema = z.object({ limits: limitsSchema.partial().optional(), timeoutMs: positive.optional() }).strict();
export const requestSchema = z.object({
  requestId: z.string().min(1), operation: z.enum(['inspect', 'execute', 'verify']),
  artifactRef: artifactSchema.optional(), options: configSchema.optional(),
  policy: policySchema.optional(), payload: z.record(z.unknown()).optional(),
}).strict();
export const engineEnvelopeSchema = z.object({
  result: z.unknown(), artifacts: z.array(artifactSchema), warnings: z.array(warningSchema),
}).strict();

export const boxSchema = z.tuple([z.number().finite(), z.number().finite(), z.number().finite(), z.number().finite()])
  .refine(([x0, y0, x1, y1]) => x1 >= x0 && y1 >= y0, 'Inverted bounding box');
export const pageSchema = z.object({
  pageNumber: positive, widthPt: z.number().finite().positive(), heightPt: z.number().finite().positive(),
  rotation: z.union([z.literal(0), z.literal(90), z.literal(180), z.literal(270)]),
  cropBox: boxSchema, userUnit: z.number().finite().positive(),
}).strict();
export const evidenceSchema = z.object({
  method: z.enum(['native', 'tagged', 'heuristic', 'ocr', 'model']),
  confidence: z.number().min(0).max(1).nullable(), engine: z.string().min(1),
}).strict();
export const physicalNodeSchema = z.object({
  pointer: z.string().min(1), pageNumber: positive,
  bbox: boxSchema.nullable(), objectRef: z.string().nullable(),
  kind: z.enum(['text', 'image', 'path', 'annotation', 'field', 'region']),
  evidence: evidenceSchema,
  // Region-level provenance alone never grants a write capability.
  editability: z.enum(['none', 'form-field', 'annotation']),
}).strict();
export const semanticNodeSchema = z.object({
  id: z.string().min(1), kind: z.enum(['heading', 'paragraph', 'table', 'figure', 'text', 'field']),
  text: z.string(), pointers: z.array(z.string().min(1)).min(1), evidence: evidenceSchema,
}).strict();
export const dualIRSchema = z.object({
  schema: z.literal('pdf-dual-ir/v1'), source: sourceSchema,
  coordinates: z.literal('crop-top-left-pt-after-rotation'),
  pages: z.array(pageSchema), physical: z.array(physicalNodeSchema), semantic: z.array(semanticNodeSchema),
  coverage: z.object({ text: z.enum(['observed', 'partial', 'unavailable']),
    readingOrder: z.enum(['observed', 'partial', 'unavailable']),
    tables: z.enum(['observed', 'partial', 'unavailable']) }).strict(),
  sourceMap: z.object({ bySemanticId: z.record(z.array(z.string())),
    byPointer: z.record(z.array(z.string())), byFingerprint: z.record(z.array(z.string())) }).strict(),
  tables: z.array(z.object({
    semanticId: z.string(), rows: positive, columns: positive,
    cells: z.array(z.object({ row: nonnegative, column: nonnegative, rowSpan: positive, columnSpan: positive,
      text: z.string(), bbox: boxSchema.nullable(), columnHeader: z.boolean(), rowHeader: z.boolean(),
      rowSection: z.boolean(), fillable: z.boolean(), pointers: z.array(z.string()) }).strict()),
  }).strict()).optional(),
}).strict();
export const parseResultSchema = z.object({ ir: dualIRSchema }).strict();
export const easyResultSchema = z.object({
  source: sourceSchema, pages: z.array(z.object({ pageNumber: positive, text: z.string(),
    coverage: z.enum(['observed', 'partial', 'unavailable']) }).strict()),
}).strict();
export const renderResultSchema = z.object({
  source: sourceSchema, engine: z.string().min(1), pageCount: positive,
  pages: z.array(z.object({ pageNumber: positive, image: artifactSchema,
    widthPx: positive, heightPx: positive }).strict()).min(1),
  visualReview: z.literal('pending'),
}).strict();
export const writtenResultSchema = z.object({
  artifactRef: artifactSchema.extend({ sha256: digestSchema, sizeBytes: positive, mediaType: z.literal('application/pdf') }),
  sources: z.array(sourceSchema), visualReview: z.literal('pending'),
}).strict();
export const deliveryResultSchema = z.object({
  manifestRef: artifactSchema.extend({ sha256: digestSchema, sizeBytes: positive }),
  document: sourceSchema, revision: positive,
  review: z.enum(['pending', 'reviewed', 'failed']),
}).strict();

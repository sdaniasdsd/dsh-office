import { z } from 'zod';
import { artifactSchema, sourceSchema, renderResultSchema, PdfModuleError } from '@dsh-office-profile/pdf-contracts';
const page = z.number().int().min(1);
const text = z.string().max(100_000);
const color = z.string().regex(/^#[0-9a-fA-F]{6}$/);
const paragraph = z.object({ kind: z.literal('paragraph'), text,
  size: z.number().min(6).max(72).default(11), color: color.default('#243447') }).strict();
const heading = z.object({ kind: z.literal('heading'), text, level: z.union([z.literal(1), z.literal(2), z.literal(3)]).default(1) }).strict();
const table = z.object({ kind: z.literal('table'), headers: z.array(text).min(1).max(12),
  rows: z.array(z.array(text)).max(5000) }).strict().refine(t => t.rows.every(r => r.length === t.headers.length), 'Table row width mismatch');
export const createPlanSchema = z.object({
  title: text.default('PDF Document'), fontRef: artifactSchema.optional(),
  pageSize: z.enum(['A4', 'LETTER']).default('A4'),
  blocks: z.array(z.union([paragraph, heading, table, z.object({ kind: z.literal('pageBreak') }).strict()])).min(1).max(10000),
}).strict();
const location = { page, x: z.number().finite().nonnegative(), y: z.number().finite().nonnegative() };
export const editPlanSchema = z.object({
  fontRef: artifactSchema.optional(),
  operations: z.array(z.discriminatedUnion('kind', [
    z.object({ kind: z.literal('rotate'), page, degrees: z.union([z.literal(90), z.literal(180), z.literal(270)]) }).strict(),
    z.object({ kind: z.literal('selectPages'), pages: z.array(page).min(1) }).strict(),
    z.object({ kind: z.literal('appendPages'), source: artifactSchema, pages: z.array(page).min(1) }).strict(),
    z.object({ kind: z.literal('stamp'), ...location, text: text.min(1), size: z.number().min(6).max(100).default(14), color: color.default('#B3261E') }).strict(),
    z.object({ kind: z.literal('addNote'), ...location, text: text.min(1) }).strict(),
    z.object({ kind: z.literal('setTextField'), name: text.min(1), value: text }).strict(),
    z.object({ kind: z.literal('setCheckbox'), name: text.min(1), checked: z.boolean() }).strict(),
  ])).min(1).max(5000),
}).strict();
export const renderPlanSchema = z.object({ dpi: z.number().int().min(36).max(300).default(120) }).strict();
export const complexPlanSchema = z.object({ ocr: z.boolean().default(true), tables: z.boolean().default(true) }).strict();
export const verifyPlanSchema = z.object({ expectedPageCount: page.optional(), textIncludes: z.array(text).optional(),
  fields: z.record(z.union([z.string(), z.boolean()])).optional(),
  notesInclude: z.array(text).optional(),
  visualReview: z.object({ source: sourceSchema, reviewedPages: z.array(page),
    findings: z.array(z.object({ page, severity: z.enum(['error','warn','info']), message: text }).strict()) }).strict().optional(),
}).strict();
export const deliveryPlanSchema = z.object({ documentId: z.string().regex(/^[A-Za-z0-9_-]{1,80}$/), revision: page,
  parentManifest: artifactSchema.optional(), preview: renderResultSchema.optional(),
  evidence: z.array(artifactSchema).max(100).default([]),
  visualReview: verifyPlanSchema.shape.visualReview,
}).strict();
export function plan<S extends z.ZodTypeAny>(schema: S, value: unknown): z.output<S> {
  const result = schema.safeParse(value ?? {});
  if (!result.success) throw new PdfModuleError('INVALID_INPUT', result.error.message);
  return result.data;
}

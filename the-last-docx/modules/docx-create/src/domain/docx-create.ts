import { z } from 'zod';
import type { ArtifactRef } from 'office-core';

export const artifactRefSchema: z.ZodType<ArtifactRef> = z.object({
  id: z.string().trim().min(1), uri: z.string().trim().min(1),
  mediaType: z.string().optional(), sizeBytes: z.number().int().nonnegative().optional(),
  sha256: z.string().regex(/^[a-f0-9]{64}$/).optional(), label: z.string().optional(),
  tags: z.record(z.string()).optional(),
}).strict();
// XML 1.0 text: no control characters or unmatched UTF-16 surrogates.
export const textSchema = z.string().max(100_000).refine(
  value => !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u.test(value),
  'Text contains invalid XML characters.',
);
const id = z.string().regex(/^[A-Za-z][A-Za-z0-9_.-]{0,79}$/);
export const presetSchema = z.enum(['report', 'technical', 'chinese-long', 'chinese-contract']);
/**
 * What the document is for. It exists so the engine can decide whether this
 * document warrants decoration at all instead of always decorating it: a formal
 * record and an academic report are both finished documents, and decorating
 * either one is a defect rather than a missing feature.
 */
export const scenarioSchema = z.enum(['internal-review', 'client-delivery', 'academic-report', 'formal-record', 'technical-spec']);
/** The resolved decoration register. `legacy` reproduces the historical output byte for byte. */
export const designRegisterSchema = z.enum(['legacy', 'report', 'academic', 'plain', 'grid']);
export const runSchema = z.object({
  text: textSchema, bold: z.boolean().optional(), italic: z.boolean().optional(),
  color: z.string().regex(/^[0-9a-fA-F]{6}$/).optional(),
}).strict();
const paragraph = z.object({
  kind: z.literal('paragraph'), id, runs: z.array(runSchema).min(1).max(1000),
  style: z.enum(['Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Caption', 'Code']).optional(),
  alignment: z.enum(['left', 'center', 'right', 'both']).optional(),
  keepWithNext: z.boolean().optional(), pageBreakBefore: z.boolean().optional(),
}).strict();
const table = z.object({
  kind: z.literal('table'), id, rows: z.array(z.array(textSchema).min(1).max(20)).min(1).max(1000),
  header: z.boolean().optional(), columnWidthsMm: z.array(z.number().positive().max(160)).optional(),
  /**
   * `register` (the default) uses the decoration register's table scheme. `none`
   * states the absence of every edge, which is what a layout table — a signature
   * block, a two-column split — needs: it is a grid of positions, not of data.
   */
  borders: z.enum(['register', 'none']).optional(),
}).strict().refine(t => t.rows.every(r => r.length === t.rows[0]!.length) &&
  (!t.columnWidthsMm || t.columnWidthsMm.length === t.rows[0]!.length), 'Table must be rectangular.');
const image = z.object({
  kind: z.literal('image'), id, artifactRef: artifactRefSchema,
  widthMm: z.number().positive().max(160), heightMm: z.number().positive().max(235),
  altText: textSchema, caption: textSchema.optional(),
}).strict();
/**
 * A real list, not hand-typed markers: the markers come from the numbering part,
 * so a renderer renumbers them and the items stay a list when the document is
 * edited later. Items are plain text the way table cells are.
 */
const list = z.object({
  kind: z.literal('list'), id,
  items: z.array(textSchema).min(1).max(1000),
  ordered: z.boolean().optional(),
}).strict();
export const blockSchema = z.union([
  paragraph, table, image, list,
  z.object({ kind: z.literal('pageBreak'), id }).strict(),
  z.object({ kind: z.literal('toc'), id, title: textSchema.optional(), levels: z.number().int().min(1).max(3).optional() }).strict(),
]);
/** Page dimensions in millimetres, the unit the rest of the plan uses. */
export const PAGE_WIDTH_MM = { A4: 210, Letter: 215.9 } as const;
export const PAGE_HEIGHT_MM = { A4: 297, Letter: 279.4 } as const;
/** The historical geometry: A4 with 25 mm all round. Omitting `page` keeps it. */
export const DEFAULT_MARGIN_MM = 25;

/**
 * Page geometry. Margins are per edge because a bound document needs a wider
 * gutter on the binding edge, and a plan that cannot say so forces the caller to
 * post-process the file.
 */
const pageSchema = z.object({
  size: z.enum(['A4', 'Letter']).default('A4'),
  marginsMm: z.object({
    top: z.number().min(0).max(100).optional(),
    right: z.number().min(0).max(100).optional(),
    bottom: z.number().min(0).max(100).optional(),
    left: z.number().min(0).max(100).optional(),
  }).strict().optional(),
}).strict();
export const documentSchema = z.object({
  preset: presetSchema.default('report'),
  scenario: scenarioSchema.optional(),
  register: designRegisterSchema.optional(),
  blocks: z.array(blockSchema).min(1).max(5000),
  header: textSchema.optional(), footer: textSchema.optional(), pageNumbers: z.boolean().default(true),
  page: pageSchema.optional(),
  /** `pageOfTotal` renders `第 X 页 共 Y 页`, the convention for bound documents. */
  pageNumberStyle: z.enum(['page', 'pageOfTotal']).default('page'),
}).strict().superRefine((doc, ctx) => {
  const ids = doc.blocks.map(b => b.id);
  if (new Set(ids).size !== ids.length) ctx.addIssue({ code: 'custom', message: 'Duplicate block IDs.' });
  if (doc.blocks.filter(b => b.kind === 'toc').length > 1) ctx.addIssue({ code: 'custom', message: 'Only one TOC is supported.' });
  // The printable width follows the margins, so a wider gutter narrows what a
  // table may occupy. Checking against a fixed 160 mm would let a table overflow
  // a page whose margins were widened.
  const width = PAGE_WIDTH_MM[doc.page?.size ?? 'A4'];
  const printable = width - (doc.page?.marginsMm?.left ?? DEFAULT_MARGIN_MM) - (doc.page?.marginsMm?.right ?? DEFAULT_MARGIN_MM);
  for (const b of doc.blocks) if (b.kind === 'table' && b.columnWidthsMm && b.columnWidthsMm.reduce((a, v) => a + v, 0) > printable + 0.5)
    ctx.addIssue({ code: 'custom', message: `Table exceeds the printable width (${Math.round(printable)} mm).` });
});
export const createPlanSchema = z.object({ kind: z.literal('create'), document: documentSchema }).strict();
export const fillPlanSchema = z.object({
  kind: z.literal('fillTemplate'), templateRef: artifactRefSchema,
  values: z.record(z.string().regex(/^[A-Za-z][A-Za-z0-9_]{0,63}$/), textSchema),
}).strict();
export const planSchema = z.union([createPlanSchema, fillPlanSchema]);
export type CreatePlan = z.infer<typeof planSchema>;
export type DocumentSpec = z.input<typeof documentSchema>;
export type Block = z.infer<typeof blockSchema>;
export type Preset = z.infer<typeof presetSchema>;
export type Scenario = z.infer<typeof scenarioSchema>;
export type DesignRegister = z.infer<typeof designRegisterSchema>;

/**
 * How much decoration a document's scenario warrants. `none` is a decision, not
 * an omission: a formal record is finished precisely because nothing was added.
 */
export type DecorationLevel = 'none' | 'restrained' | 'moderate';

/**
 * The resolved decoration decision. It travels back with the result so a
 * reviewer can disagree with the judgement itself rather than only its output.
 */
export interface DesignDecision {
  register: DesignRegister;
  label: string;
  decoration: DecorationLevel;
  intent: string;
  /** How the decision was reached. */
  source: 'explicit' | 'scenario' | 'compatibility-default';
  reason: string;
  scenario?: Scenario;
}

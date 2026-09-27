import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import { configSchema } from './config';
import { artifactRefSchema, designRegisterSchema, planSchema, scenarioSchema } from './domain/docx-create';
import { failure, ERROR_CODES } from './errors';
const base = {
  requestId: z.string().trim().min(1).max(200), options: configSchema.optional(),
  policy: z.object({ id: z.string().min(1), allowMacros: z.boolean().optional(), allowExternalLinks: z.boolean().optional(),
    allowEncrypted: z.boolean().optional(), maxBlocks: z.number().int().positive().optional(), maxTableCells: z.number().int().positive().optional() }).strict().optional(),
};
export const executeSchema = z.object({ ...base, operation: z.literal('execute'), plan: planSchema }).strict();
export const inspectSchema = z.object({ ...base, operation: z.literal('inspect'), artifactRef: artifactRefSchema }).strict();
export const verifySchema = z.object({ ...base, operation: z.literal('verify'), artifactRef: artifactRefSchema }).strict();
const reportSchema = z.object({
  ok: z.boolean(), partial: z.boolean(),
  checks: z.array(z.object({ id: z.string(), status: z.enum(['pass', 'fail', 'skip']), severity: z.enum(['info', 'warn', 'error']), message: z.string() }).strict()),
  summary: z.object({ total: z.number().int().nonnegative(), passed: z.number().int().nonnegative(), failed: z.number().int().nonnegative(), skipped: z.number().int().nonnegative() }).strict(),
}).strict();
export const executeOutputSchema = z.object({
  moduleId: z.literal('docx-create'), requestId: z.string(), operation: z.literal('execute'),
  result: z.object({ artifactRef: artifactRefSchema, engine: z.string(), mode: z.enum(['create', 'fillTemplate']), paragraphIds: z.array(z.string()),
    filledKeys: z.array(z.string()), fields: z.enum(['pending-update', 'none']), visualReview: z.literal('pending'),
    // Omitted for `fillTemplate`, where the template owns its own styling.
    design: z.object({
      register: designRegisterSchema, label: z.string(), decoration: z.enum(['none', 'restrained', 'moderate']),
      intent: z.string(), source: z.enum(['explicit', 'scenario', 'compatibility-default']), reason: z.string(),
      scenario: scenarioSchema.optional(),
    }).strict().optional() }).strict(),
  artifacts: z.array(artifactRefSchema),
  warnings: z.array(z.object({ code: z.string(), message: z.string(), severity: z.enum(['info', 'warn', 'error']) }).strict()),
  verification: reportSchema.optional(),
}).strict();
export const errorSchema = z.object({ name: z.literal('DocxCreateError'), code: z.enum(ERROR_CODES), message: z.string() }).strict();
export const DOCX_CREATE_SCHEMAS = {
  config: zodToJsonSchema(configSchema, { $refStrategy: 'none' }),
  inspect: zodToJsonSchema(inspectSchema, { $refStrategy: 'none' }),
  execute: zodToJsonSchema(executeSchema, { $refStrategy: 'none' }),
  verify: zodToJsonSchema(verifySchema, { $refStrategy: 'none' }),
  executeOutput: zodToJsonSchema(executeOutputSchema, { $refStrategy: 'none' }),
  verificationReport: zodToJsonSchema(reportSchema, { $refStrategy: 'none' }),
  error: zodToJsonSchema(errorSchema, { $refStrategy: 'none' }),
};
/** JSON.stringify silently drops functions, NaN and undefined; reject lossy input before schema parsing. */
export function assertSerializable(value: unknown, seen = new Set<object>(), depth = 0): void {
  if (depth > 50) failure('INVALID_INPUT', 'Input nesting is too deep.');
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number' && Number.isFinite(value)) return;
  if (typeof value !== 'object' || seen.has(value)) failure('INVALID_INPUT', 'Input must be a plain JSON value.');
  if (!Array.isArray(value) && Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null) failure('INVALID_INPUT', 'Input must use plain objects.');
  seen.add(value);
  for (const child of Object.values(value)) assertSerializable(child, seen, depth + 1);
  seen.delete(value);
}
export function parseInput<S extends z.ZodTypeAny>(schema: S, input: unknown): z.output<S> {
  assertSerializable(input);
  const result = schema.safeParse(input);
  if (!result.success) failure('INVALID_INPUT', result.error.issues.map(i => `${i.path.join('.')}: ${i.message}`).join('; '));
  return result.data;
}

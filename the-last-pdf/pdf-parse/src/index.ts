import { createPdfModule, parseResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_PARSE_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-parse', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfParseModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_PARSE_DEFINITION, parseResultSchema, withDefaultEngine('pdf-parse', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfParseModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfParseModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

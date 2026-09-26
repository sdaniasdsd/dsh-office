import { createPdfModule, easyResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_EASY_PARSE_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-easy-parse', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfEasyParseModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_EASY_PARSE_DEFINITION, easyResultSchema, withDefaultEngine('pdf-easy-parse', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfEasyParseModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfEasyParseModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

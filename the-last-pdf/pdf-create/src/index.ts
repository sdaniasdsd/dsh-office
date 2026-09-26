import { createPdfModule, writtenResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_CREATE_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-create', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfCreateModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_CREATE_DEFINITION, writtenResultSchema, withDefaultEngine('pdf-create', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfCreateModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfCreateModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

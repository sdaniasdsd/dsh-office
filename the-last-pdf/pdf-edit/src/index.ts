import { createPdfModule, writtenResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_EDIT_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-edit', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfEditModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_EDIT_DEFINITION, writtenResultSchema, withDefaultEngine('pdf-edit', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfEditModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfEditModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

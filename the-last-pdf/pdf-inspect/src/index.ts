import { createPdfModule, profileSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_INSPECT_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-inspect', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfInspectModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_INSPECT_DEFINITION, profileSchema, withDefaultEngine('pdf-inspect', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfInspectModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfInspectModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

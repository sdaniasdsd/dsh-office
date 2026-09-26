import { createPdfModule, renderResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_RENDER_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-render', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfRenderModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_RENDER_DEFINITION, renderResultSchema, withDefaultEngine('pdf-render', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfRenderModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfRenderModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

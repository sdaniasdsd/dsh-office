import { createPdfModule, deliveryResultSchema } from '@dsh-office-profile/pdf-contracts';
import type { ModuleDefinition, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import definition from '../module.json';
import { withDefaultEngine } from '@dsh-office-profile/pdf-engines';
export const PDF_ARTIFACT_DEFINITION = Object.freeze({ ...definition,
  id: 'pdf-artifact', profileGroup: 'PDF', capabilities: ['inspect', 'execute', 'verify'] as const,
  implementation: 'implemented', engineSelection: 'user-approved',
} satisfies ModuleDefinition);
export function createPdfArtifactModule(options: ModuleOptions = {}) {
  return createPdfModule(PDF_ARTIFACT_DEFINITION, deliveryResultSchema, withDefaultEngine('pdf-artifact', options));
}
export async function register(registry: { registerModule(module: ReturnType<typeof createPdfArtifactModule>): void | Promise<void> }, options: ModuleOptions = {}) {
  const module = createPdfArtifactModule(options);
  await registry.registerModule(module);
  return module;
}
export * from './contract';

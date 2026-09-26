export { DocxProfile,MODULE_IDS } from './profile';
export type { ProfileOptions,ModuleId,ProfileRegistryLike,RegisteredDocxModule } from './profile';
export { LocalArtifactFiles } from './files';
export { bridgeComplexToDual } from './source-bridge';
export { doctor } from './doctor';
export { createMcpServer } from './mcp';
// Preserve the entire public surface, including dual-IR revision ledger/governance/brief.
export * as docxInspect from '@dsh-office-profile/docx-inspect';
export * as docxEasyParse from '@dsh-office-profile/docx-easy-parse';
export * as docxParse from '@dsh-office-profile/docx-parse';
export * as docxComplexParse from '@dsh-office-profile/docx-complex-parse';
export * as docxCreate from '@dsh-office-profile/docx-create';
export * as docxEdit from '@dsh-office-profile/docx-edit';
export * as docxRender from '@dsh-office-profile/docx-render';
export * as docxArtifact from '@dsh-office-profile/docx-artifact';

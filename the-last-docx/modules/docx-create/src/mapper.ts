import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { CreateResult } from './contract';
import type { PackageFacts } from './engine/package-check';
import { DOCX_MIME } from './engine/package-check';
export function toFormatProfile(facts: PackageFacts): FormatProfile {
  return { format: 'docx', mediaType: DOCX_MIME, extension: 'docx', confidence: 1,
    container: 'zip', encrypted: false, signatures: ['zip:checked', 'content-types:docx', 'relationships:checked'],
    features: { fields: facts.fields, editableParagraphIds: facts.paragraphIds.length > 0 },
    metadata: { paragraphs: facts.paragraphs, tables: facts.tables, entries: facts.entries } };
}
export function toCreateResult(artifactRef: ArtifactRef, engine: string, mode: CreateResult['mode'], filledKeys: string[], facts: PackageFacts): CreateResult {
  return { artifactRef, engine, mode, filledKeys: [...filledKeys], paragraphIds: facts.paragraphIds,
    fields: facts.fields ? 'pending-update' : 'none', visualReview: 'pending' };
}
export function pendingWarnings(fields: boolean): Warning[] {
  return [
    { code: 'VISUAL_REVIEW_PENDING', message: '尚未检查最终分页、重叠和截断，请交给 docx-render。', severity: 'info' },
    { code: 'FONT_ENVIRONMENT_DEPENDENT', message: '未嵌入字体；接收端字体缺失可能改变换行和分页。', severity: 'warn' },
    ...(fields ? [{ code: 'FIELDS_PENDING_UPDATE', message: '目录和页码字段尚未由排版程序更新。', severity: 'warn' as const }] : []),
  ];
}

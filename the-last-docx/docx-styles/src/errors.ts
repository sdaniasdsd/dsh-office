import type { DocxStylesErrorCode } from './contract';

export class DocxStylesError extends Error {
  readonly code: DocxStylesErrorCode;
  readonly details?: Record<string, string | number | boolean>;

  constructor(code: DocxStylesErrorCode, message: string, details?: Record<string, string | number | boolean>) {
    super(message);
    this.name = 'DocxStylesError';
    this.code = code;
    if (details) this.details = details;
  }
}

export function toDocxStylesError(error: unknown): DocxStylesError {
  if (error instanceof DocxStylesError) return error;
  return new DocxStylesError('ENGINE_FAILED', error instanceof Error ? error.message : 'DOCX style authoring failed');
}

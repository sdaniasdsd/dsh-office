import type { DocxEditErrorCode } from './contract';

export class DocxEditError extends Error {
  readonly code: DocxEditErrorCode;
  readonly details?: Record<string, string | number | boolean>;

  constructor(code: DocxEditErrorCode, message: string, details?: Record<string, string | number | boolean>) {
    super(message);
    this.name = 'DocxEditError';
    this.code = code;
    if (details) this.details = details;
  }
}

export function toDocxEditError(error: unknown): DocxEditError {
  if (error instanceof DocxEditError) return error;
  return new DocxEditError('ENGINE_FAILED', error instanceof Error ? error.message : 'DOCX edit failed');
}

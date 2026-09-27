import type { DocxRenderErrorCode } from './contract';

export class DocxRenderError extends Error {
  readonly code: DocxRenderErrorCode;
  readonly details?: Record<string, string | number | boolean>;
  constructor(code: DocxRenderErrorCode, message: string, details?: Record<string, string | number | boolean>) {
    super(message);
    this.name = 'DocxRenderError';
    this.code = code;
    if (details) this.details = details;
  }
}
export function toDocxRenderError(error: unknown): DocxRenderError {
  if (error instanceof DocxRenderError) return error;
  const message = error instanceof Error ? error.message : 'DOCX render failed.';
  if (/ENOENT|not found|cannot find/i.test(message)) return new DocxRenderError('ENGINE_UNAVAILABLE', message);
  return new DocxRenderError('ENGINE_FAILED', message);
}

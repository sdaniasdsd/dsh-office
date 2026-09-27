export const ERROR_CODES = [
  'INVALID_INPUT', 'FORMAT_MISMATCH', 'SAFETY_POLICY_DENIED', 'LIMIT_EXCEEDED',
  'ARTIFACT_STORE_UNAVAILABLE', 'ARTIFACT_INTEGRITY_FAILED', 'ENGINE_FAILED',
  'ENGINE_TIMEOUT', 'VERIFICATION_FAILED', 'TEMPLATE_VALUE_MISSING',
  'TEMPLATE_VALUE_UNUSED', 'UNSUPPORTED_TEMPLATE', 'MODULE_DISPOSED',
] as const;
export type ErrorCode = typeof ERROR_CODES[number];
export class DocxCreateError extends Error {
  constructor(readonly code: ErrorCode, message: string) {
    super(message); this.name = 'DocxCreateError';
  }
  toJSON() { return { name: this.name, code: this.code, message: this.message }; }
}
export function failure(code: ErrorCode, message: string): never { throw new DocxCreateError(code, message); }
export function normalizeError(error: unknown): DocxCreateError {
  return error instanceof DocxCreateError ? error : new DocxCreateError('ENGINE_FAILED', 'DOCX engine failed.');
}

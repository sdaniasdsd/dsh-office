export const ERROR_CODES = ['INVALID_INPUT','FORMAT_MISMATCH','SAFETY_POLICY_DENIED','LIMIT_EXCEEDED','ARTIFACT_UNAVAILABLE','INTEGRITY_MISMATCH','STALE_REFERENCE','VERSION_CONFLICT','ENGINE_FAILED','ENGINE_TIMEOUT','VERIFICATION_FAILED','MODULE_DISPOSED'] as const;
export type ErrorCode = typeof ERROR_CODES[number];
export class DocxArtifactError extends Error {
  constructor(readonly code: ErrorCode, message: string) { super(message); this.name = 'DocxArtifactError'; }
  toJSON() { return { name: this.name, code: this.code, message: this.message }; }
}
export function fail(code: ErrorCode, message: string): never { throw new DocxArtifactError(code, message); }
export function normalized(error: unknown): DocxArtifactError {
  return error instanceof DocxArtifactError ? error : new DocxArtifactError('ENGINE_FAILED', 'Delivery adapter failed.');
}

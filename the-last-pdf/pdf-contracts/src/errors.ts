export const ERROR_CODES = [
  'INVALID_INPUT', 'ARTIFACT_NOT_FOUND', 'ARTIFACT_MISMATCH',
  'ENGINE_UNAVAILABLE', 'ENGINE_FAILED', 'ENGINE_PROTOCOL_ERROR', 'ENGINE_TIMEOUT',
  'LIMIT_EXCEEDED', 'SAFETY_POLICY_DENIED', 'SOURCE_MAP_MISMATCH',
  'MODULE_DISPOSED',
  'FORMAT_MISMATCH', 'PASSWORD_REQUIRED', 'UNSUPPORTED_OPERATION', 'FONT_REQUIRED',
  'ARTIFACT_CONFLICT', 'VERIFICATION_FAILED',
] as const;
export type ErrorCode = typeof ERROR_CODES[number];

export class PdfModuleError extends Error {
  constructor(public readonly code: ErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'PdfModuleError';
  }
  toJSON() { return { name: this.name, code: this.code, message: this.message }; }
}

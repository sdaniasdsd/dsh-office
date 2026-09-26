/** Standalone development boundary only. Field-compatible with docx-parse's
 * ArtifactRef / Warning. Replace with real office-* types at Profile integration.
 * Do not extend the DOCX-only FormatKind union via a conflicting declaration. */
declare module 'office-core' {
  export interface ArtifactRef {
    id: string;
    uri: string;
    mediaType?: string;
    sizeBytes?: number;
    sha256?: string;
    label?: string;
    tags?: Record<string, string>;
  }
  export interface Warning {
    code: string;
    message: string;
    severity: 'info' | 'warn' | 'error';
    path?: string;
    details?: Record<string, unknown>;
  }
}
declare module 'office-test-kit' {
  export interface VerificationReportLike {
    ok: boolean;
    partial: boolean;
    checks: unknown[];
    summary: { total: number; passed: number; failed: number; skipped: number };
  }
}

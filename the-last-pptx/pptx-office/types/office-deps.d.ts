/** Standalone module stubs; the real Profile supplies these contracts at integration. */
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

declare module 'office-safety' {
  export interface SafetyPolicy {
    id: string;
    allowMacros?: boolean;
    allowExternalLinks?: boolean;
    allowEncrypted?: boolean;
  }
}

/** Standalone-only boundary stubs; real Profile packages supply these contracts at integration. */
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
  export interface FormatProfile {
    format: 'docx' | 'docm' | 'dotx' | 'dotm' | 'ole' | 'rtf' | 'zip' | 'unknown';
    mediaType: string | null;
    extension: string | null;
    confidence: number;
    container: 'zip' | 'ole' | 'rtf' | 'unknown';
    encrypted: boolean | null;
    signatures: string[];
    features?: Record<string, boolean | null>;
    metadata?: Record<string, string | number | boolean | null>;
  }
  export interface FormatIR {
    profile: FormatProfile;
    parts: unknown[];
    relationships: unknown[];
    indicators: { macros: unknown[]; externalReferences: unknown[]; embeddedObjects: unknown[] };
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

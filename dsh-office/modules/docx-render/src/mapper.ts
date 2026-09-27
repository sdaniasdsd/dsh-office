import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type { VisualFinding } from './contract';

export function toFormatProfile(
  artifactRef: ArtifactRef,
  inspected: { extension: string | null; mediaType: string | null },
): FormatProfile {
  const format = inspected.extension === '.docm' ? 'docm' : inspected.extension === '.dotx' ? 'dotx' : inspected.extension === '.dotm' ? 'dotm' : 'docx';
  return {
    format,
    mediaType: inspected.mediaType,
    extension: inspected.extension?.slice(1) ?? null,
    // The module validates the ZIP signature and declared extension; full OOXML
    // package validation remains the parser/inspector's responsibility.
    confidence: 0.65,
    container: 'zip',
    encrypted: null,
    signatures: ['zip:local-header', `declared-extension:${inspected.extension ?? 'unknown'}`],
    features: {},
    metadata: { byteLength: artifactRef.sizeBytes ?? 0 },
  };
}

export function findingsToWarnings(findings: VisualFinding[]): Warning[] {
  return findings.map((finding) => ({
    code: 'VISUAL_FINDINGS_PRESENT',
    severity: finding.severity,
    message: finding.message,
    path: finding.sourcePointers?.[0],
    details: {
      findingId: finding.id,
      kind: finding.kind,
      page: finding.page,
      ...(finding.confidence === undefined ? {} : { confidence: finding.confidence }),
      ...(finding.sourceSemanticIds?.length ? { sourceSemanticIds: finding.sourceSemanticIds } : {}),
    },
  }));
}

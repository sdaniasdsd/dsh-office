import { createDualIR, sha256 } from '@dsh-office-profile/pdf-contracts';
import type { PdfEngine, PdfFormatProfile } from '@dsh-office-profile/pdf-contracts';
export const bytes = new TextEncoder().encode('synthetic bytes; intentionally not a PDF');
export const ref = { id: 'fixture', uri: 'memory:fixture', sha256: sha256(bytes), sizeBytes: bytes.length };
export const source = { id: ref.id, sha256: ref.sha256 };
export const reader = { async read() { return bytes.slice(); } };
export const profile: PdfFormatProfile = {
  format: 'pdf', mediaType: 'application/pdf', container: 'pdf', version: '1.7', pageCount: 1,
  encrypted: null, features: { javascript: null, embeddedFiles: false, externalLinks: false,
    acroForm: false, xfa: false, signatures: null, tagged: false }, coverage: 'partial',
};
export function engine(result: unknown): PdfEngine {
  return { name: 'synthetic-test-only', async invoke() { return { result, artifacts: [], warnings: [] }; } };
}
export function irFixture() {
  const evidence = { method: 'native' as const, confidence: null, engine: 'synthetic-test-only' };
  return createDualIR({
    schema: 'pdf-dual-ir/v1', source, coordinates: 'crop-top-left-pt-after-rotation',
    pages: [{ pageNumber: 1, widthPt: 100, heightPt: 200, cropBox: [0, 0, 100, 200], rotation: 0, userUnit: 1 }],
    physical: [{ pointer: 'page/1/text/0', pageNumber: 1, bbox: [10, 20, 30, 40], objectRef: null,
      kind: 'text', editability: 'none', evidence }],
    semantic: [{ id: 'semantic-1', kind: 'text', text: '重复文字', pointers: ['page/1/text/0'], evidence }],
    coverage: { text: 'observed', readingOrder: 'partial', tables: 'unavailable' },
  });
}

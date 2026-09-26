import { createPdfParseModule } from '../pdf-parse/src/index';
import { createDualIR, sha256 } from '@dsh-office-profile/pdf-contracts';

// Architecture demo only. These bytes are deliberately NOT a real PDF.
const bytes = new TextEncoder().encode('synthetic-framework-demo');
const artifactRef = { id: 'demo', uri: 'memory:demo', sha256: sha256(bytes) };
const module = createPdfParseModule({
  files: { async read() { return bytes.slice(); } },
  engine: { name: 'synthetic-demo-only', async invoke(context) {
    return { result: { ir: createDualIR({
      schema: 'pdf-dual-ir/v1', source: context.source!.identity,
      coordinates: 'crop-top-left-pt-after-rotation',
      pages: [{ pageNumber: 1, widthPt: 595, heightPt: 842, rotation: 0, cropBox: [0, 0, 595, 842], userUnit: 1 }],
      physical: [], semantic: [],
      coverage: { text: 'unavailable', readingOrder: 'unavailable', tables: 'unavailable' },
    }) }, artifacts: [], warnings: [{ code: 'SYNTHETIC_DEMO', severity: 'info', message: 'Framework demonstration; no PDF was parsed.' }] };
  } },
});
try {
  console.log(JSON.stringify(await module.handlers.execute({ artifactRef, requestId: 'demo', operation: 'execute' }), null, 2));
} finally { await module.dispose(); }

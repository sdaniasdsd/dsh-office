import { PdfModuleError, verifyPolicy } from '@dsh-office-profile/pdf-contracts';
import type { ModuleOptions, ModuleId, PdfEngine } from '@dsh-office-profile/pdf-contracts';
import { inspectPdf, parsePdf, renderPdf, verifyPdf } from './native';
import { createPdf, editPdf } from './write';
import { deliver, verifyDelivery } from './delivery';
import { complexPdf } from './complex';
export * from './store';
export * from './plans';

export function createDefaultEngine(id: ModuleId, options: ModuleOptions): PdfEngine {
  return { name: 'pdfjs-qpdf-pdfkit-pdflib-docling', async invoke(context) {
    if (context.request.operation === 'verify') return id === 'pdf-artifact' ? verifyDelivery(context, options) : verifyPdf(context, options);
    if (context.request.operation === 'inspect' || id === 'pdf-inspect') return { result: await inspectPdf(context), artifacts: [], warnings: [] };
    if (id === 'pdf-create') return createPdf(context, options);
    if (id === 'pdf-edit') return editPdf(context, options);
    if (id === 'pdf-render') return renderPdf(context, options);
    if (id === 'pdf-artifact') return deliver(context, options);
    if (id === 'pdf-complex-parse') return complexPdf(context, options);
    const ir = await parsePdf(context);
    const warnings = [{ code: 'NATIVE_TEXT_ONLY', severity: 'info', message: 'Native text extracted; reading order, OCR and table semantics are not inferred.' }];
    if (id === 'pdf-easy-parse') {
      const pointerPages = new Map(ir.physical.map(p => [p.pointer, p.pageNumber]));
      const texts = new Map<number, string[]>();
      for (const node of ir.semantic) for (const page of new Set(node.pointers.map(p => pointerPages.get(p)!))) {
        const fragments = texts.get(page) ?? []; fragments.push(node.text); texts.set(page, fragments);
      }
      return { result: { source: ir.source, pages: ir.pages.map(page => ({ pageNumber: page.pageNumber,
        text: (texts.get(page.pageNumber) ?? []).join(''), coverage: 'partial' })) }, artifacts: [], warnings };
    }
    return { result: { ir }, artifacts: [], warnings };
  } };
}
export function withDefaultEngine(id: ModuleId, options: ModuleOptions): ModuleOptions {
  if (options.engine || !options.files) return options;
  const engine = createDefaultEngine(id, options);
  return { ...options, engine, safetyGuard: options.safetyGuard ?? { async assertAllowed(context) {
    // artifact.verify reads a JSON manifest; document policy belongs to delivery-time inspection.
    if (id === 'pdf-artifact' && context.request.operation === 'verify') throw new PdfModuleError('INVALID_INPUT', 'Apply PDF policy to document inspection, not the manifest.');
    if (!context.source) return;
    if (!verifyPolicy(await inspectPdf(context), context.policy).ok) throw new PdfModuleError('SAFETY_POLICY_DENIED', 'PDF does not satisfy the declared policy.');
  } } };
}

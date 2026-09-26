import { fileURLToPath } from 'node:url';
import { z } from 'zod';
import { createDualIR, PdfModuleError } from '@dsh-office-profile/pdf-contracts';
import type { ArtifactRef, EngineContext, ModuleOptions, PdfDualIR } from '@dsh-office-profile/pdf-contracts';
import { plan, complexPlanSchema } from './plans';
import { runProcess, withPdfFile } from './process';
import { withDocument } from './native';

const box = z.object({ l: z.number().finite(), t: z.number().finite(), r: z.number().finite(), b: z.number().finite(), coord_origin: z.enum(['TOPLEFT','BOTTOMLEFT']) });
const observation = z.object({ protocol: z.literal('docling-observations/v1'), version: z.string(), backend: z.literal('pypdfium2'), status: z.string(),
  pages: z.record(z.object({ size: z.object({ width: z.number().positive(), height: z.number().positive() }) })),
  items: z.array(z.object({ self_ref: z.string(), label: z.string(), text: z.string().optional(),
    prov: z.array(z.object({ page_no: z.number().int().positive(), bbox: box })),
      data: z.object({ num_rows: z.number().int().nonnegative(), num_cols: z.number().int().nonnegative(),
      table_cells: z.array(z.object({ text: z.string(), bbox: box.nullable().optional(), column_header: z.boolean().optional(),
        row_header: z.boolean().optional(), row_section: z.boolean().optional(), fillable: z.boolean().optional(), start_row_offset_idx: z.number().int().nonnegative(),
        start_col_offset_idx: z.number().int().nonnegative(), row_span: z.number().int().positive(), col_span: z.number().int().positive() })) }).optional(),
  })),
});
export async function complexPdf(context: EngineContext, options: ModuleOptions) {
  const input = plan(complexPlanSchema, context.request.payload);
  if (!options.backend?.modelPath) throw new PdfModuleError('ENGINE_UNAVAILABLE', 'Docling requires an explicitly provisioned modelPath.');
  const pages = await withDocument(context, async doc => {
    const pages: PdfDualIR['pages'] = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const p = await doc.getPage(n), v = p.getViewport({ scale: 1 });
      pages.push({ pageNumber: n, widthPt: v.width, heightPt: v.height, rotation: p.rotate as 0|90|180|270,
        cropBox: [0, 0, v.width, v.height], userUnit: p.userUnit });
    }
    return pages;
  });
  const result = await withPdfFile(context.source!.bytes, path => runProcess(options.backend?.pythonPath ?? 'python',
    [options.backend?.doclingScript ?? fileURLToPath(new URL('./docling_bridge.py', import.meta.url)), path,
      '--models', options.backend!.modelPath!, '--ocr', String(input.ocr), '--tables', String(input.tables),
      '--max-pages', String(context.config.limits.maxPages), '--max-bytes', String(context.config.limits.maxInputBytes)],
    context.signal, context.config.limits.maxOutputJsonBytes));
  if (result.code !== 0) throw new PdfModuleError('ENGINE_FAILED', `Docling failed: ${result.stderr.slice(-1500)}`);
  let raw: z.infer<typeof observation>;
  let fullObservation: unknown;
  try { fullObservation = JSON.parse(result.stdout); raw = observation.parse(fullObservation); }
  catch { throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Docling returned invalid observations.'); }
  if (raw.status !== 'success') throw new PdfModuleError('ENGINE_FAILED', `Docling conversion status: ${raw.status}`);
  const physical: PdfDualIR['physical'] = [], semantic: PdfDualIR['semantic'] = [], tables: NonNullable<PdfDualIR['tables']> = [];
  let mappedCells = 0;
  const evidence = { method: 'model' as const, confidence: null, engine: `docling-${raw.version}/${raw.backend}` };
  for (const item of raw.items) {
    if (physical.length + semantic.length + mappedCells + (item.data?.table_cells.length ?? 0) + item.prov.length + 1 > context.config.limits.maxNodes)
      throw new PdfModuleError('LIMIT_EXCEEDED', 'Docling observations exceed node budget.');
    if (!item.prov.length) continue;
    const id = `${context.source!.identity.sha256.slice(0,16)}:${item.self_ref}`, pointers: string[] = [];
    for (const [i, prov] of item.prov.entries()) {
      const page = pages[prov.page_no - 1], modelPage = raw.pages[String(prov.page_no)];
      if (!page || !modelPage) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Docling provenance references missing page.');
      const pointer = `${item.self_ref}/prov/${i}`, b = prov.bbox;
      // Only claim geometry where model and native coordinate systems agree.
      const compatible = page.rotation === 0 && page.userUnit === 1 && page.cropBox[0] === 0 && page.cropBox[1] === 0 &&
        Math.abs(page.widthPt - modelPage.size.width) < 0.5 && Math.abs(page.heightPt - modelPage.size.height) < 0.5;
      const y1 = b.coord_origin === 'BOTTOMLEFT' ? modelPage.size.height - b.t : b.t;
      const y2 = b.coord_origin === 'BOTTOMLEFT' ? modelPage.size.height - b.b : b.b;
      physical.push({ pointer, pageNumber: prov.page_no, bbox: compatible ? [Math.min(b.l,b.r), Math.min(y1,y2), Math.max(b.l,b.r), Math.max(y1,y2)] : null,
        objectRef: null, kind: 'region', evidence, editability: 'none' });
      pointers.push(pointer);
    }
    const isTable = item.label === 'table' && item.data && item.data.num_rows > 0 && item.data.num_cols > 0;
    if (isTable && item.data) {
      mappedCells += item.data.table_cells.length;
      const pageNo = item.prov.length === 1 ? item.prov[0]!.page_no : null;
      const page = pageNo === null ? null : pages[pageNo - 1]!;
      const modelPage = pageNo === null ? null : raw.pages[String(pageNo)]!;
      const compatible = page !== null && modelPage !== null && page.rotation === 0 && page.userUnit === 1 &&
        Math.abs(page.widthPt - modelPage.size.width) < 0.5 && Math.abs(page.heightPt - modelPage.size.height) < 0.5;
      const cells: NonNullable<PdfDualIR['tables']>[number]['cells'] = item.data.table_cells.map((c, cellIndex) => {
        const pointer = `${item.self_ref}/cell/${cellIndex}`;
        let bbox: [number, number, number, number] | null = null;
        if (compatible && pageNo !== null && modelPage && c.bbox) {
          const y1 = c.bbox.coord_origin === 'BOTTOMLEFT' ? modelPage.size.height - c.bbox.t : c.bbox.t;
          const y2 = c.bbox.coord_origin === 'BOTTOMLEFT' ? modelPage.size.height - c.bbox.b : c.bbox.b;
          bbox = [Math.min(c.bbox.l,c.bbox.r), Math.min(y1,y2), Math.max(c.bbox.l,c.bbox.r), Math.max(y1,y2)];
          physical.push({ pointer, pageNumber: pageNo, bbox, objectRef: null, kind: 'region', evidence, editability: 'none' });
        }
        return { row: c.start_row_offset_idx, column: c.start_col_offset_idx, rowSpan: c.row_span,
          columnSpan: c.col_span, text: c.text, bbox, columnHeader: c.column_header ?? false,
          rowHeader: c.row_header ?? false, rowSection: c.row_section ?? false, fillable: c.fillable ?? false,
          pointers: bbox ? [pointer] : [] };
      });
      tables.push({ semanticId: id, rows: item.data.num_rows, columns: item.data.num_cols, cells });
      pointers.push(...cells.flatMap(c => c.pointers));
    }
    semantic.push({ id, kind: isTable ? 'table' : item.label === 'section_header' || item.label === 'title' ? 'heading' : item.label === 'picture' ? 'figure' : 'paragraph',
      text: item.text ?? item.data?.table_cells.map(c => c.text).join('\t') ?? '', pointers, evidence });
  }
  const ir = createDualIR({ schema: 'pdf-dual-ir/v1', source: context.source!.identity,
    coordinates: 'crop-top-left-pt-after-rotation', pages, physical, semantic, tables,
    coverage: { text: 'partial', readingOrder: 'partial', tables: input.tables ? 'partial' : 'unavailable' } });
  const artifacts: ArtifactRef[] = [];
  if (options.writer) {
    const bytes = Buffer.from(JSON.stringify({ schema: 'pdf-evidence/v1', source: context.source!.identity, payload: fullObservation }));
    if (bytes.length > context.config.limits.maxOutputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Raw observations exceed output budget.');
    artifacts.push(await options.writer.write({ bytes, sources: [context.source!.ref], suggestedName: 'docling-observations.json', requestId: context.request.requestId, signal: context.signal }));
  }
  return { result: { ir }, artifacts, warnings: [{ code: 'MODEL_OBSERVATIONS', severity: 'info',
    message: 'Model-derived structure, not edit authority. Table cells preserve upstream text, spans, roles and compatible per-cell boxes in IR; all upstream fields remain in the full source-bound observation. Rotated/multi-page ambiguous table geometry stays null; confidence is not calibrated.' }] };
}

import type { z } from 'zod';
import { dualIRSchema } from './schema';
import type { PdfDualIR } from './contract';
import { PdfModuleError } from './errors';
import { sha256 } from './validation';

type IRBody = Omit<z.infer<typeof dualIRSchema>, 'sourceMap'>;
export function buildSourceMap(semantic: PdfDualIR['semantic']): PdfDualIR['sourceMap'] {
  const bySemanticId: Record<string, string[]> = Object.create(null);
  const byPointer: Record<string, string[]> = Object.create(null);
  const byFingerprint: Record<string, string[]> = Object.create(null);
  for (const node of semantic) {
    if (Object.hasOwn(bySemanticId, node.id)) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Duplicate semantic ID.');
    bySemanticId[node.id] = [...node.pointers];
    for (const pointer of node.pointers) (byPointer[pointer] ??= []).push(node.id);
    (byFingerprint[sha256(node.text)] ??= []).push(node.id);
  }
  return { bySemanticId, byPointer, byFingerprint };
}
export function createDualIR(body: IRBody): PdfDualIR {
  return validateDualIR({ ...body, sourceMap: buildSourceMap(body.semantic) });
}
export function validateDualIR(value: unknown): PdfDualIR {
  const parsed = dualIRSchema.safeParse(value);
  if (!parsed.success) throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', parsed.error.message);
  const ir = parsed.data;
  if (ir.pages.length === 0 || ir.pages.some((p, i) => p.pageNumber !== i + 1))
    throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'IR pages must be complete and one-based.');
  const pointers = new Set<string>();
  for (const node of ir.physical) {
    if (pointers.has(node.pointer) || node.pageNumber > ir.pages.length)
      throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Duplicate pointer or nonexistent page.');
    pointers.add(node.pointer);
    if ((node.editability === 'form-field' && node.kind !== 'field') ||
        (node.editability === 'annotation' && node.kind !== 'annotation') ||
        (node.editability !== 'none' && node.objectRef === null))
      throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Writable targets require a field/annotation object reference.');
    if ((node.evidence.method === 'ocr' || node.evidence.method === 'model') && node.editability !== 'none')
      throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Model/OCR observations do not grant edit authority.');
  }
  const physicalByPointer = new Map(ir.physical.map(node => [node.pointer, node]));
  for (const node of ir.semantic) {
    if (new Set(node.pointers).size !== node.pointers.length || node.pointers.some(p => !pointers.has(p)))
      throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Semantic node references missing or duplicate pointers.');
  }
  const tableIds = new Set<string>();
  for (const table of ir.tables ?? []) {
    if (tableIds.has(table.semanticId) || !ir.semantic.some(n => n.id === table.semanticId && n.kind === 'table'))
      throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Table requires a unique table semantic node.');
    tableIds.add(table.semanticId);
    for (const cell of table.cells) {
      const locations = cell.pointers.map(pointer => physicalByPointer.get(pointer));
      if (cell.row + cell.rowSpan > table.rows || cell.column + cell.columnSpan > table.columns ||
        new Set(cell.pointers).size !== cell.pointers.length || cell.pointers.some(p => !pointers.has(p)) ||
        (cell.bbox === null ? cell.pointers.length > 0 : cell.pointers.length === 0 ||
          locations.some(node => !node || node.kind !== 'region' || JSON.stringify(node.bbox) !== JSON.stringify(cell.bbox))))
        throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Invalid table cell span or pointer.');
    }
  }
  const expected = buildSourceMap(ir.semantic);
  for (const index of ['bySemanticId', 'byPointer', 'byFingerprint'] as const) {
    const actual = ir.sourceMap[index];
    if (Object.keys(actual).length !== Object.keys(expected[index]).length ||
        Object.entries(expected[index]).some(([key, ids]) => !Object.hasOwn(actual, key) ||
          JSON.stringify([...actual[key]!].sort()) !== JSON.stringify([...ids].sort())))
      throw new PdfModuleError('SOURCE_MAP_MISMATCH', `Invalid source map: ${index}.`);
  }
  return ir;
}
export function assertSameSource(actual: { id: string; sha256: string }, expected: { id: string; sha256: string }): void {
  if (actual.id !== expected.id || actual.sha256 !== expected.sha256)
    throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Evidence belongs to a different artifact revision.');
}
/** Affine transform supplied by the chosen engine, including CropBox translation,
 * UserUnit, rotation, and origin conversion. Transform all corners, not just two.
 * Result is an axis-aligned envelope, not a glyph polygon or writable selector. */
export function transformBox(
  box: readonly [number, number, number, number],
  matrix: readonly [number, number, number, number, number, number],
): [number, number, number, number] {
  if (![...box, ...matrix].every(Number.isFinite) || box[2] < box[0] || box[3] < box[1])
    throw new PdfModuleError('INVALID_INPUT', 'Invalid box or affine transform.');
  const [a, b, c, d, e, f] = matrix;
  if (a * d - b * c === 0) throw new PdfModuleError('INVALID_INPUT', 'Singular transform.');
  const corners = [[box[0], box[1]], [box[2], box[1]], [box[0], box[3]], [box[2], box[3]]] as const;
  const xs = corners.map(([x, y]) => a * x + c * y + e);
  const ys = corners.map(([x, y]) => b * x + d * y + f);
  const result: [number, number, number, number] = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
  if (!result.every(Number.isFinite)) throw new PdfModuleError('INVALID_INPUT', 'Coordinate overflow.');
  return result;
}

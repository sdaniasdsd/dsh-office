import { describe, it, expect } from 'vitest';
import { createDualIR, validateDualIR, transformBox, sha256 } from '@dsh-office-profile/pdf-contracts';
import { irFixture } from './support';

describe('PDF provenance and geometry', () => {
  it('supports many-to-many provenance and duplicate text without losing candidates', () => {
    const ir = irFixture();
    ir.semantic.push({ ...ir.semantic[0]!, id: 'semantic-2' });
    const result = createDualIR(ir);
    expect(result.sourceMap.byPointer['page/1/text/0']).toEqual(['semantic-1', 'semantic-2']);
    expect(result.sourceMap.byFingerprint[sha256('重复文字')]).toEqual(['semantic-1', 'semantic-2']);
  });
  it('detects stale reverse indexes', () => {
    const ir = irFixture();
    ir.sourceMap.byPointer['page/1/text/0'] = ['missing'];
    expect(() => validateDualIR(ir)).toThrow('Invalid source map');
  });
  it('rejects dangling pointers and duplicate semantic IDs', () => {
    const ir = irFixture();
    ir.semantic[0]!.pointers = ['nonexistent'];
    expect(() => createDualIR(ir)).toThrow('missing');
    const another = irFixture();
    another.semantic.push(another.semantic[0]!);
    expect(() => createDualIR(another)).toThrow('Duplicate');
  });
  it('does not grant write authority to OCR or generic regions', () => {
    const ir = irFixture();
    Object.assign(ir.physical[0]!, { kind: 'field', editability: 'form-field', objectRef: '9 0 R' });
    ir.physical[0]!.evidence.method = 'ocr';
    expect(() => validateDualIR(ir)).toThrow('edit authority');
    ir.physical[0]!.evidence.method = 'native';
    ir.physical[0]!.kind = 'region';
    expect(() => validateDualIR(ir)).toThrow('field/annotation');
  });
  it('keeps absent geometry null', () => {
    const ir = irFixture();
    ir.physical[0]!.bbox = null;
    expect(validateDualIR(ir).physical[0]!.bbox).toBeNull();
  });
  it('requires table cell geometry to match its physical source region', () => {
    const base = irFixture();
    const ir = createDualIR({ schema: base.schema, source: base.source, coordinates: base.coordinates,
      pages: base.pages, physical: [{ ...base.physical[0]!, pointer: 'table/cell/0', kind: 'region' }],
      semantic: [{ id: 'table-1', kind: 'table', text: 'cell', pointers: ['table/cell/0'], evidence: base.semantic[0]!.evidence }],
      tables: [{ semanticId: 'table-1', rows: 1, columns: 1, cells: [{ row: 0, column: 0, rowSpan: 1, columnSpan: 1,
        text: 'cell', bbox: [10,20,30,40], columnHeader: false, rowHeader: false, rowSection: false, fillable: false, pointers: ['table/cell/0'] }] }],
      coverage: { ...base.coverage, tables: 'observed' } });
    expect(ir.tables?.[0]?.cells[0]?.pointers).toEqual(['table/cell/0']);
    const invalid = structuredClone(ir);
    invalid.tables![0]!.cells[0]!.pointers = [];
    expect(() => validateDualIR(invalid)).toThrow(/table cell/);
  });
  it.each([
    [[1, 0, 0, -1, 0, 200], [10, 160, 30, 180]],
    [[0, 1, 1, 0, 0, 0], [20, 10, 40, 30]],
    [[-1, 0, 0, 1, 100, 0], [70, 20, 90, 40]],
    [[0, -1, -1, 0, 200, 100], [160, 70, 180, 90]],
    [[2, 0, 0, -2, -10, 400], [10, 320, 50, 360]],
  ])('normalizes rotation/crop/user-unit via adapter affine transform %#', (matrix, expected) => {
    expect(transformBox([10, 20, 30, 40], matrix as [number, number, number, number, number, number])).toEqual(expected);
  });
  it('rejects singular transforms and coordinate overflow', () => {
    expect(() => transformBox([0, 0, 1, 1], [0, 0, 0, 0, 0, 0])).toThrow('Singular');
    expect(() => transformBox([0, 0, 2, 2], [1e308, 0, 0, 1e308, 0, 0])).toThrow('overflow');
  });
});

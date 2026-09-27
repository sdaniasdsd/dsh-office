import { createHash } from 'node:crypto';
import ExcelJS from 'exceljs';
import { canonicalJson } from '@dsh-office-profile/docx-artifact';
import { describe, expect, it } from 'vitest';
import { createXlsxOfficeModule } from '../src/index';

function store() {
  const objects = new Map<string, Uint8Array>();
  return {
    objects,
    async write(input: { bytes: Uint8Array; suggestedName: string }) {
      const hash = createHash('sha256').update(input.bytes).digest('hex');
      const ref = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/${input.suggestedName}`, sha256: hash, sizeBytes: input.bytes.length, label: input.suggestedName, mediaType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
      objects.set(ref.uri, input.bytes); return ref;
    },
    async read(ref: { uri: string }) { const bytes = objects.get(ref.uri); if (!bytes) throw new Error('missing'); return bytes; },
  };
}

describe('xlsx-office', () => {
  it('creates immutable XLSX, reopens it, and verifies expected worksheet', async () => {
    const artifacts = store(), module = createXlsxOfficeModule({ artifactStore: artifacts, pythonPath: 'python' });
    const created = await module.handlers.execute({ requestId: 'create-test', operation: 'execute', payload: { action: 'createWorkbook', sheetName: 'Data', rows: [['Name', 'Amount'], ['Ada', 12]] } }) as { result: { artifactRef: { uri: string } } };
    expect(artifacts.objects.has(created.result.artifactRef.uri)).toBe(true);
    const verified = await module.handlers.verify({ requestId: 'verify-test', operation: 'verify', artifactRef: created.result.artifactRef, payload: { expectedWorksheetCount: 1, sheetNamesInclude: ['Data'] } }) as { result: { ok: boolean } };
    expect(verified.result.ok).toBe(true);
  });

  it('reads bounded cells and writes only literal values to a new artifact', async () => {
    const artifacts = store(), wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 'before'; ws.getCell('B1').value = { formula: '1+1', result: 2 };
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const sourceHash = createHash('sha256').update(bytes).digest('hex');
    const source = { id: `sha256:${sourceHash}`, uri: 'file:///managed/source/source.xlsx', sha256: sourceHash, sizeBytes: bytes.length, label: 'source.xlsx' };
    artifacts.objects.set(source.uri, bytes);
    const module = createXlsxOfficeModule({ artifactStore: artifacts, pythonPath: 'python' });
    const read = await module.handlers.execute({ requestId: 'read-test', operation: 'execute', artifactRef: source, payload: { action: 'readRange', sheet: 'Data', range: 'A1:B1' } }) as { result: { cells: { value: unknown }[] } };
    expect(() => canonicalJson(read)).not.toThrow();
    expect(read.result.cells.map(cell => cell.value)).toEqual(['before', { formula: '1+1', sharedFormula: null, result: 2 }]);
    const edited = await module.handlers.execute({ requestId: 'write-test', operation: 'execute', artifactRef: source, payload: { action: 'setCells', changes: [{ sheet: 'Data', address: 'A1', value: '=not-a-formula' }] } }) as { result: { artifactRef: { uri: string } } };
    const out = new ExcelJS.Workbook(); await out.xlsx.load(Buffer.from(artifacts.objects.get(edited.result.artifactRef.uri)!) as unknown as Parameters<typeof out.xlsx.load>[0]);
    expect(out.getWorksheet('Data')!.getCell('A1').value).toBe('=not-a-formula');
  });
});

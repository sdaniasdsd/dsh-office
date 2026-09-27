import { createHash } from 'node:crypto';
import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
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

  it('changes only selected worksheet print setup and preserves other OOXML parts', async () => {
    const artifacts = store(), wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 'kept';
    wb.addWorksheet('Other').getCell('A1').value = 'untouched';
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const sourceHash = createHash('sha256').update(bytes).digest('hex');
    const source = { id: `sha256:${sourceHash}`, uri: 'file:///managed/print-source/source.xlsx', sha256: sourceHash, sizeBytes: bytes.length, label: 'source.xlsx' };
    artifacts.objects.set(source.uri, bytes);
    const module = createXlsxOfficeModule({ artifactStore: artifacts, pythonPath: 'python' });
    const output = await module.handlers.execute({ requestId: 'print-layout-test', operation: 'execute', artifactRef: source, payload: {
      action: 'setPrintLayout', sheets: [{ sheet: 'Data', orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 }],
    } }) as { result: { artifactRef: { uri: string }; packageScopedWrite: boolean; untouchedPackagePartsPreservedByteForByte: boolean; sheetContentPreserved: boolean; changedPackageParts: { part: string }[] } };
    expect(output.result.packageScopedWrite).toBe(true);
    expect(output.result.untouchedPackagePartsPreservedByteForByte).toBe(true);
    expect(output.result.sheetContentPreserved).toBe(true);
    expect(output.result.changedPackageParts.map((item) => item.part)).toEqual(['xl/worksheets/sheet1.xml']);
    const withoutDirectories = { filter: (entry: { name: string }) => !entry.name.endsWith('/') };
    const before = unzipSync(bytes, withoutDirectories), after = unzipSync(artifacts.objects.get(output.result.artifactRef.uri)!, withoutDirectories);
    for (const [part, original] of Object.entries(before)) {
      if (part === 'xl/worksheets/sheet1.xml') continue;
      expect(after[part]).toEqual(original);
    }
    const sheetXml = strFromU8(after['xl/worksheets/sheet1.xml']!);
    expect(sheetXml).toContain('<pageSetUpPr fitToPage="1"/>');
    expect(sheetXml).toMatch(/<pageSetup\b[^>]*orientation="landscape"[^>]*fitToWidth="1"[^>]*fitToHeight="0"[^>]*\/>/);
    expect(sheetXml).not.toContain('scale=');
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(Buffer.from(artifacts.objects.get(output.result.artifactRef.uri)!) as unknown as Parameters<typeof out.xlsx.load>[0]);
    expect(out.getWorksheet('Data')!.getCell('A1').value).toBe('kept');
    expect(out.getWorksheet('Other')!.getCell('A1').value).toBe('untouched');
  });

  it('formats bounded header ranges while preserving cell values and formulas', async () => {
    const artifacts = store(), wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 'Quarter'; ws.getCell('B1').value = 'Revenue'; ws.getCell('A2').value = 'Q1';
    ws.getCell('B2').value = { formula: 'SUM(1,2)', result: 3 };
    const other = wb.addWorksheet('Other'); other.getCell('A1').value = 'Keep';
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const hash = createHash('sha256').update(bytes).digest('hex');
    const source = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/source.xlsx`, sha256: hash, sizeBytes: bytes.length, label: 'source.xlsx' };
    artifacts.objects.set(source.uri, bytes);
    const module = createXlsxOfficeModule({ artifactStore: artifacts, pythonPath: 'python' });
    const edited = await module.handlers.execute({ requestId: 'format-cells-test', operation: 'execute', artifactRef: source, payload: {
      action: 'formatCells', changes: [{ sheet: '*', range: 'A1:Z1', font: { bold: true, color: '#FFFFFF' }, fill: '#244A67', border: 'thin' }],
    } }) as { result: { artifactRef: { uri: string }; formattedCells: number } };
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(Buffer.from(artifacts.objects.get(edited.result.artifactRef.uri)!) as unknown as Parameters<typeof out.xlsx.load>[0]);
    expect(out.getWorksheet('Data')!.getCell('A1').value).toBe('Quarter');
    expect(out.getWorksheet('Data')!.getCell('B2').value).toMatchObject({ formula: 'SUM(1,2)', result: 3 });
    expect(out.getWorksheet('Data')!.getCell('A1').font.bold).toBe(true);
    expect(out.getWorksheet('Data')!.getCell('A1').fill).toMatchObject({ fgColor: { argb: 'FF244A67' } });
    expect(out.getWorksheet('Other')!.getCell('A1').value).toBe('Keep');
    expect(edited.result.formattedCells).toBe(3);
  });

  it('rejects empty styles and overlapping format ranges', async () => {
    const artifacts = store(), wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 'Name'; ws.getCell('B1').value = 'Amount'; ws.getCell('C1').value = 'Status';
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer()), hash = createHash('sha256').update(bytes).digest('hex');
    const source = { id: `sha256:${hash}`, uri: `file:///managed/${hash}/overlap.xlsx`, sha256: hash, sizeBytes: bytes.length, label: 'overlap.xlsx' };
    artifacts.objects.set(source.uri, bytes);
    const module = createXlsxOfficeModule({ artifactStore: artifacts, pythonPath: 'python' });
    await expect(module.handlers.execute({ requestId: 'empty-style', operation: 'execute', artifactRef: source, payload: {
      action: 'formatCells', changes: [{ sheet: 'Data', range: 'A1:B1', font: {} }],
    } })).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    await expect(module.handlers.execute({ requestId: 'overlap-style', operation: 'execute', artifactRef: source, payload: {
      action: 'formatCells', changes: [
        { sheet: 'Data', range: 'A1:B1', font: { bold: true } },
        { sheet: 'Data', range: 'B1:C1', fill: '#244A67' },
      ],
    } })).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });
});

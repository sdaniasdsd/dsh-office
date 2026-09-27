import ExcelJS from 'exceljs';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { ArtifactRef, Warning } from 'office-core';
import type { SafetyPolicy } from 'office-safety';
import { z } from 'zod';

export const XLSX_OFFICE_MODULE_ID = 'xlsx-office' as const;
export const XLSX_OFFICE_CAPABILITIES = ['inspect', 'execute', 'verify'] as const;
export const XLSX_OFFICE_DEFINITION = Object.freeze({
  id: XLSX_OFFICE_MODULE_ID, version: '0.1.0', profileGroup: 'XLSX', capabilities: XLSX_OFFICE_CAPABILITIES,
  summary: 'Inspect XLSX workbooks, read ranges, write guarded literal cells, and create workbooks with ExcelJS.',
  dependencies: [{ name: 'exceljs', kind: 'runtime' }, { name: 'office-core', kind: 'module' }, { name: 'office-files', kind: 'module' }, { name: 'office-safety', kind: 'module' }],
  configSchema: { type: 'object', properties: { engine: { type: 'object', properties: { pythonPath: { type: 'string', default: 'python' } } } } },
});
const LIMITS = Object.freeze({ maxInputBytes: 64 * 1024 * 1024, maxOutputBytes: 64 * 1024 * 1024,
  maxSheets: 200, maxRows: 100_000, maxColumns: 512, maxCells: 250_000, maxChanges: 20_000, maxTextChars: 2_000_000 });
const artifactSchema = z.object({ id: z.string().min(1), uri: z.string().min(1), sha256: z.string().regex(/^[a-f0-9]{64}$/).optional(),
  sizeBytes: z.number().int().nonnegative().optional(), mediaType: z.string().optional(), label: z.string().optional() }).passthrough();
const literalSchema = z.union([z.string().max(100_000), z.number().finite(), z.boolean(), z.null()]);
const cellChangeSchema = z.object({ sheet: z.string().min(1).max(255), address: z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/i), value: literalSchema }).strict();
const requestSchema = z.object({ requestId: z.string().min(1), operation: z.enum(XLSX_OFFICE_CAPABILITIES), artifactRef: artifactSchema.optional(),
  policy: z.object({ id: z.string().min(1) }).passthrough().optional(), payload: z.record(z.unknown()).optional() }).passthrough();
const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

export interface XlsxArtifactStore {
  read(ref: ArtifactRef, maxBytes?: number, signal?: AbortSignal): Promise<Uint8Array>;
  write(input: { bytes: Uint8Array; suggestedName: string; requestId?: string; source?: ArtifactRef; sources?: ArtifactRef[] }): Promise<ArtifactRef>;
}
export interface XlsxOfficeOptions { artifactStore: XlsxArtifactStore; pythonPath?: string;
  safetyGuard?: { assertAllowed(input: { artifactRef: ArtifactRef; operation: string; policy: SafetyPolicy }): Promise<void> } }
class XlsxOfficeError extends Error { constructor(readonly code: string, message: string) { super(message); this.name = 'XlsxOfficeError'; } }

function assertRef(ref: ArtifactRef) {
  if (extname(ref.label ?? ref.uri.split(/[?#]/, 1)[0] ?? '').toLowerCase() !== '.xlsx') throw new XlsxOfficeError('FORMAT_MISMATCH', 'xlsx-office accepts macro-free .xlsx files only.');
}
function safeCell(value: ExcelJS.CellValue): unknown {
  if (value === null || value === undefined || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value ?? null;
  if (value instanceof Date) return { date: value.toISOString() };
  if (Array.isArray(value)) return value.map(item => safeCell(item as ExcelJS.CellValue));
  if (typeof value === 'object') {
    const record = value as unknown as Record<string, unknown>;
    if (typeof record.formula === 'string' || typeof record.sharedFormula === 'string') return {
      formula: record.formula ?? null, sharedFormula: record.sharedFormula ?? null,
      result: record.result === undefined ? null : safeCell(record.result as ExcelJS.CellValue),
    };
    if (Array.isArray(record.richText)) return { richText: record.richText.map((part: { text?: string }) => String(part.text ?? '')).join('') };
    if (typeof record.text === 'string' && typeof record.hyperlink === 'string') return { text: record.text, hyperlink: record.hyperlink };
    if (typeof record.error === 'string') return { error: record.error };
    return JSON.parse(JSON.stringify(value));
  }
  return null;
}
function workbookSummary(workbook: ExcelJS.Workbook) {
  if (workbook.worksheets.length > LIMITS.maxSheets) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Worksheet count exceeds the Profile limit.');
  let cells = 0, workbookTextChars = 0;
  const worksheets = workbook.worksheets.map(sheet => {
    if (sheet.rowCount > LIMITS.maxRows || sheet.columnCount > LIMITS.maxColumns) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Worksheet dimensions exceed the Profile limit.');
    let formulas = 0, sheetCells = 0;
    sheet.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
      cells++; sheetCells++;
      if (cell.type === ExcelJS.ValueType.Formula) formulas++;
      workbookTextChars += typeof cell.value === 'string' ? cell.value.length : JSON.stringify(cell.value ?? '').length;
    }));
    if (cells > LIMITS.maxCells) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Workbook populated-cell count exceeds the Profile limit.');
    return { name: sheet.name, rowCount: sheet.rowCount, columnCount: sheet.columnCount, populatedCells: sheetCells, formulaCount: formulas, state: sheet.state };
  });
  if (workbookTextChars > LIMITS.maxTextChars) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Workbook text exceeds the Profile character limit.');
  return { worksheetCount: worksheets.length, populatedCellCount: cells, worksheets };
}
async function loadWorkbook(bytes: Uint8Array) {
  const workbook = new ExcelJS.Workbook();
  const copy = Buffer.alloc(bytes.length); copy.set(bytes);
  await workbook.xlsx.load(copy as unknown as Parameters<typeof workbook.xlsx.load>[0]);
  workbookSummary(workbook);
  return workbook;
}
function runProbe(command: string, script: string, inputPath: string, timeoutMs: number) {
  return new Promise<{ code: number; stdout: string; stderr: string }>((resolve, reject) => {
    const child = spawn(command, [script, inputPath], { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '', stderr = '', settled = false;
    const finish = (error?: unknown, value?: { code: number; stdout: string; stderr: string }) => { if (settled) return; settled = true; clearTimeout(timer); error ? reject(error) : resolve(value!); };
    const timer = setTimeout(() => { child.kill(); finish(new XlsxOfficeError('ENGINE_TIMEOUT', 'XLSX package preflight exceeded the time limit.')); }, timeoutMs);
    child.stdout.on('data', (chunk: Buffer) => { stdout += chunk.toString('utf8'); if (stdout.length > 4096) { child.kill(); finish(new XlsxOfficeError('LIMIT_EXCEEDED', 'XLSX preflight response is too large.')); } });
    child.stderr.on('data', (chunk: Buffer) => { stderr += chunk.toString('utf8'); if (stderr.length > 8192) { child.kill(); finish(new XlsxOfficeError('LIMIT_EXCEEDED', 'XLSX preflight diagnostics are too large.')); } });
    child.once('error', () => finish(new XlsxOfficeError('ENGINE_UNAVAILABLE', 'Could not start the configured Python runtime for XLSX ZIP preflight.')));
    child.once('close', code => finish(undefined, { code: code ?? -1, stdout, stderr }));
  });
}

export function createXlsxOfficeModule(options: XlsxOfficeOptions) {
  if (!options?.artifactStore) throw new XlsxOfficeError('INVALID_INPUT', 'artifactStore is required.');
  let disposed = false;
  const invoke = async (inputRaw: unknown, operation: 'inspect' | 'execute' | 'verify') => {
    const parsed = requestSchema.safeParse(inputRaw);
    if (!parsed.success || parsed.data.operation !== operation) throw new XlsxOfficeError('INVALID_INPUT', parsed.success ? 'Handler operation mismatch.' : parsed.error.message);
    const input = parsed.data, payload = input.payload ?? {};
    if (disposed) throw new XlsxOfficeError('MODULE_DISPOSED', 'XLSX module has been disposed.');
    const action = operation === 'inspect' ? 'inspect' : operation === 'verify' ? 'verify' : String(payload.action ?? '');
    if (operation === 'execute' && !['readRange', 'setCells', 'createWorkbook'].includes(action)) throw new XlsxOfficeError('INVALID_INPUT', 'execute payload.action must be readRange, setCells, or createWorkbook.');
    const create = action === 'createWorkbook';
    if (!create && !input.artifactRef) throw new XlsxOfficeError('INVALID_INPUT', 'artifactRef is required for this XLSX operation.');
    if (input.artifactRef) assertRef(input.artifactRef);
    if (input.policy) {
      if (!input.artifactRef || !options.safetyGuard) throw new XlsxOfficeError('SAFETY_POLICY_DENIED', 'A safety policy requires an artifactRef and configured guard.');
      await options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, operation, policy: input.policy as SafetyPolicy });
    }
    let bytes: Uint8Array<ArrayBufferLike> = new Uint8Array(0);
    let digest: string | undefined;
    let unsupportedFeatures: string[] = [];
    if (input.artifactRef) {
      try { bytes = await options.artifactStore.read(input.artifactRef, LIMITS.maxInputBytes); }
      catch { throw new XlsxOfficeError('ARTIFACT_NOT_FOUND', 'XLSX artifact could not be read from managed storage.'); }
      if (bytes.length > LIMITS.maxInputBytes) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'XLSX input exceeds byte budget.');
      digest = sha256(bytes);
      if ((input.artifactRef.sha256 && input.artifactRef.sha256 !== digest) || (input.artifactRef.sizeBytes !== undefined && input.artifactRef.sizeBytes !== bytes.length)) throw new XlsxOfficeError('ARTIFACT_INTEGRITY_FAILED', 'XLSX artifact metadata does not match its bytes.');
      if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new XlsxOfficeError('FORMAT_MISMATCH', 'XLSX must be a ZIP-based OOXML package.');
    }
    const temp = await mkdtemp(join(tmpdir(), 'dsh-xlsx-'));
    try {
      const inputPath = join(temp, 'input.xlsx');
      if (input.artifactRef) {
        await writeFile(inputPath, bytes);
        const probe = await runProbe(options.pythonPath ?? 'python', fileURLToPath(new URL('./engine/zip_probe.py', import.meta.url)), inputPath, 15_000);
        if (probe.code !== 0) {
          if (/No module named/.test(probe.stderr)) throw new XlsxOfficeError('ENGINE_UNAVAILABLE', 'Python ZIP preflight runtime is unavailable.');
          throw new XlsxOfficeError('UNSUPPORTED_OPERATION', probe.stderr.trim().slice(-1000));
        }
        try { unsupportedFeatures = z.object({ unsupportedFeatures: z.array(z.string()) }).parse(JSON.parse(probe.stdout)).unsupportedFeatures; }
        catch { throw new XlsxOfficeError('ENGINE_PROTOCOL_ERROR', 'XLSX preflight returned invalid feature metadata.'); }
        if (action === 'setCells' && unsupportedFeatures.length) throw new XlsxOfficeError('UNSUPPORTED_OPERATION', `ExcelJS may not preserve workbook features on write: ${unsupportedFeatures.join(', ')}.`);
      }
      let workbook: ExcelJS.Workbook;
      if (create) {
        const sheetName = z.string().min(1).max(31).safeParse(payload.sheetName ?? 'Sheet1');
        const rows = z.array(z.array(literalSchema).max(LIMITS.maxColumns)).max(LIMITS.maxRows).safeParse(payload.rows ?? []);
        if (!sheetName.success || !rows.success || rows.data.reduce<number>((n, row) => n + row.length, 0) > LIMITS.maxCells || rows.success && rows.data.reduce<number>((n, row) => n + row.reduce<number>((chars, value) => chars + (typeof value === 'string' ? value.length : 0), 0), 0) > LIMITS.maxTextChars) throw new XlsxOfficeError('INVALID_INPUT', 'createWorkbook requires a valid sheetName and bounded literal rows.');
        workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet(sheetName.data);
        for (const row of rows.data) sheet.addRow(row);
      } else workbook = await loadWorkbook(bytes);
      const summary = workbookSummary(workbook);
      const featureWarnings: Warning[] = unsupportedFeatures.map(feature => ({ code: 'UNSUPPORTED_WORKBOOK_FEATURE', severity: 'warn', message: `Workbook contains ${feature}; ExcelJS may not preserve it if saved. Cell edits are denied.` }));
      if (operation === 'inspect') return { moduleId: XLSX_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: { format: 'xlsx', ...summary, unsupportedFeatures, engine: 'ExcelJS 4.4.0' }, artifacts: [], warnings: [...featureWarnings, { code: 'FORMULA_CACHE_NOT_CALCULATED', severity: 'info', message: 'ExcelJS reads formula text and cached values but does not calculate formulas.' }] as Warning[] };
      if (operation === 'verify') {
        const expects = z.object({ expectedWorksheetCount: z.number().int().nonnegative().optional(), sheetNamesInclude: z.array(z.string()).optional() }).passthrough().safeParse(payload);
        if (!expects.success) throw new XlsxOfficeError('INVALID_INPUT', expects.error.message);
        const checks = [
          { id: 'xlsx.reopened', status: 'pass' as const, severity: 'error' as const, message: 'Workbook loaded and bounded structure extracted.' },
          ...(expects.data.expectedWorksheetCount === undefined ? [] : [{ id: 'xlsx.worksheetCount', status: expects.data.expectedWorksheetCount === workbook.worksheets.length ? 'pass' as const : 'fail' as const, severity: 'error' as const, message: 'Worksheet count compared with expectation.' }]),
          ...(expects.data.sheetNamesInclude ?? []).map((name, index) => ({ id: `xlsx.sheet.${index}`, status: workbook.getWorksheet(name) ? 'pass' as const : 'fail' as const, severity: 'error' as const, message: `Worksheet ${name} existence compared with expectation.` })),
        ];
        const report = { ok: checks.every(item => item.status !== 'fail'), checks, summary: { total: checks.length, failed: checks.filter(item => item.status === 'fail').length } };
        return { moduleId: XLSX_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: report, verification: report, artifacts: [], warnings: [] as Warning[] };
      }
      if (action === 'readRange') {
        const request = z.object({ action: z.literal('readRange'), sheet: z.string().min(1), range: z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?$/i) }).strict().safeParse(payload);
        if (!request.success) throw new XlsxOfficeError('INVALID_INPUT', request.error.message);
        const sheet = workbook.getWorksheet(request.data.sheet); if (!sheet) throw new XlsxOfficeError('INVALID_INPUT', 'Worksheet does not exist.');
        const bounds = request.data.range.toUpperCase().split(':');
        const decode = (address: string) => { const match = /^([A-Z]+)([1-9]\d*)$/.exec(address)!; let col = 0; for (const char of match[1]!) col = col * 26 + char.charCodeAt(0) - 64; return { row: Number(match[2]), col }; };
        const start = decode(bounds[0]!), end = decode(bounds[1] ?? bounds[0]!);
        if (start.row > LIMITS.maxRows || end.row > LIMITS.maxRows || start.col > LIMITS.maxColumns || end.col > LIMITS.maxColumns || end.row < start.row || end.col < start.col) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Requested range is reversed or outside worksheet limits.');
        const cells: unknown[] = []; let count = 0;
        for (let row = start.row; row <= end.row; row++) for (let col = start.col; col <= end.col; col++) {
          if (++count > LIMITS.maxCells) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Requested range exceeds the cell limit.');
          const cell = sheet.getCell(row, col);
          const formula = cell.type === ExcelJS.ValueType.Formula && typeof cell.value === 'object' && cell.value ? (cell.value as ExcelJS.CellFormulaValue).formula : undefined;
          cells.push({ address: cell.address, value: safeCell(cell.value), ...(formula ? { formula } : {}) });
        }
        return { moduleId: XLSX_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: { sheet: sheet.name, range: request.data.range, cells, unsupportedFeatures, formulaPolicy: 'cached-values-only-no-recalculation' }, artifacts: [], warnings: [...featureWarnings, { code: 'FORMULA_CACHE_NOT_CALCULATED', severity: 'info', message: 'Formula cached values are reported as-is; ExcelJS does not recalculate formulas.' }] as Warning[] };
      }
      if (action === 'setCells') {
        const request = z.object({ action: z.literal('setCells'), changes: z.array(cellChangeSchema).min(1).max(LIMITS.maxChanges) }).strict().safeParse(payload);
        if (!request.success) throw new XlsxOfficeError('INVALID_INPUT', request.error.message);
        const seen = new Set<string>();
        let textChars = 0;
        for (const change of request.data.changes) {
          const match = /^([A-Z]+)([1-9]\d*)$/i.exec(change.address)!; let col = 0; for (const char of match[1]!.toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
          if (Number(match[2]) > LIMITS.maxRows || col > LIMITS.maxColumns) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Cell address exceeds worksheet limits.');
          if (typeof change.value === 'string') textChars += change.value.length;
          if (textChars > LIMITS.maxTextChars) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Cell values exceed the Profile character limit.');
          const key = `${change.sheet}!${change.address.toUpperCase()}`; if (seen.has(key)) throw new XlsxOfficeError('INVALID_INPUT', 'A cell may only be changed once per request.'); seen.add(key);
          const sheet = workbook.getWorksheet(change.sheet); if (!sheet) throw new XlsxOfficeError('INVALID_INPUT', `Worksheet ${change.sheet} does not exist.`);
          sheet.getCell(change.address).value = change.value;
        }
      }
      const outputBytes = new Uint8Array(await workbook.xlsx.writeBuffer());
      if (!outputBytes.length || outputBytes.length > LIMITS.maxOutputBytes) throw new XlsxOfficeError('LIMIT_EXCEEDED', 'Generated XLSX is empty or exceeds the output byte budget.');
      const reopened = await loadWorkbook(outputBytes);
      const artifactRef = await options.artifactStore.write({ bytes: outputBytes, requestId: input.requestId, ...(input.artifactRef ? { sources: [{ ...input.artifactRef, sha256: digest, sizeBytes: bytes.length }] } : {}), suggestedName: create ? 'created.xlsx' : 'edited.xlsx' });
      return { moduleId: XLSX_OFFICE_MODULE_ID, requestId: input.requestId, operation, result: { artifactRef, sources: input.artifactRef ? [{ id: input.artifactRef.id, sha256: digest }] : [], ...workbookSummary(reopened), formulaPolicy: 'cached-values-only-no-recalculation', visualReview: 'pending' }, artifacts: [artifactRef], warnings: [...featureWarnings, { code: 'VISUAL_REVIEW_PENDING', severity: 'info', message: 'Workbook was reopened and structurally checked; review in Excel or LibreOffice when visual appearance matters.' }, { code: 'FORMULA_CACHE_NOT_CALCULATED', severity: 'info', message: 'ExcelJS does not calculate formulas; formula caches may be stale.' }] as Warning[] };
    } finally { await rm(temp, { recursive: true, force: true }); }
  };
  return { definition: XLSX_OFFICE_DEFINITION, handlers: { inspect: (input: unknown) => invoke(input, 'inspect'), execute: (input: unknown) => invoke(input, 'execute'), verify: (input: unknown) => invoke(input, 'verify') }, async dispose() { disposed = true; } };
}
export async function register(registry: { registerModule(module: ReturnType<typeof createXlsxOfficeModule>): void | Promise<void> }, options: XlsxOfficeOptions) {
  const module = createXlsxOfficeModule(options); await registry.registerModule(module); return module;
}

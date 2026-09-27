import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { createDocxComplexParseModule } from '../src/index';
import { parseRawResult, type RawParseResult } from '../src/domain/docx-complex-parse';
import { toComplexContent, toComplexIR, toFormatProfile } from '../src/mapper';
import { verify } from '../src/verifier';
import { paragraph, rawResult } from './support';

const python = process.env['DOCX_PARSE_PYTHON'] ?? 'python';
const engineScript = resolve('src/engine/docx_complex_parse.py');
let temporary: string;
let path: string;

function runEngine(flags: Record<string, unknown> = {}, limits = {}) {
  const process = spawnSync(python, [engineScript, '--path', path, '--config', JSON.stringify({
    featureFlags: { parsePageGeometry: false, ...flags }, limits,
  })], { encoding: 'utf8', env: { ...globalThis.process.env, PYTHONIOENCODING: 'utf-8', PYTHONDONTWRITEBYTECODE: '1' } });
  expect(process.error).toBeUndefined();
  expect(process.status, process.stderr).toBe(0);
  return JSON.parse(process.stdout);
}

function irFor(raw: RawParseResult) {
  return toComplexIR(raw, toComplexContent(raw).content, toFormatProfile(raw, { extension: 'docx' }));
}

function referenceCheck(raw: RawParseResult, mutate: (ir: ReturnType<typeof irFor>) => void) {
  const ir = irFor(raw);
  mutate(ir);
  return verify({ ir, confidenceFloor: 0.5 }).checks.find((item) => item.id === 'content.references');
}

beforeAll(() => {
  temporary = mkdtempSync(join(tmpdir(), 'docx-complex-test-'));
  path = join(temporary, '双IR.docx');
  const result = spawnSync(python, [resolve('tests/create_fixture.py'), path], { encoding: 'utf8' });
  expect(result.error).toBeUndefined();
  expect(result.status, result.stderr).toBe(0);
});
afterAll(() => { if (temporary) rmSync(temporary, { recursive: true, force: true }); });

describe('real ZIP → Python observations → validated domain → dual IR', () => {
  it('handles continuous sections, revisions, nested breaks and alternate drawing branches', () => {
    const result = spawnSync(python, ['-B', resolve('tests/test_observations.py')], { encoding: 'utf8' });
    expect(result.error).toBeUndefined();
    expect(result.status, result.stderr).toBe(0);
  });
  it('preserves break position without inventing a current page number', () => {
    const raw = parseRawResult(runEngine());
    const ir = irFor(raw);
    expect(ir.content.breaks.map((item) => item.kind)).toEqual(['section', 'explicit', 'explicit', 'rendered']);
    expect(ir.content.breaks[0]?.beforeNodeId).toBe(ir.content.blocks[1]?.id);
    expect(ir.content.breaks[1]?.beforeNodeId).toBe(ir.content.blocks[1]?.id);
    expect(ir.content.breaks.slice(2).every((item) => item.beforeNodeId === null)).toBe(true);
    expect(ir.content.breaks.every((item) => item.pageIndexAfter === null && item.coordinate.box === null)).toBe(true);
  });

  it('reads section end properties into the correct start boundary and preserves units', () => {
    const content = irFor(parseRawResult(runEngine())).content;
    expect(content.sections).toHaveLength(2);
    expect(content.sections[0]?.startBeforeNodeId).toBe(content.blocks[0]?.id);
    expect(content.sections[1]).toMatchObject({ startBeforeNodeId: content.blocks[1]?.id,
      columnCount: 2, startType: 'oddPage', geometry: { width: 15840, height: 12240, unit: 'twip', orientation: 'landscape' } });
  });

  it('joins DrawingML to package relationships, retaining negative relative offsets', () => {
    const raw = parseRawResult(runEngine());
    const ir = irFor(raw);
    expect(ir.content.floats).toHaveLength(2);
    expect(ir.content.floats[0]).toMatchObject({ kind: 'image', relationshipId: 'rImg', name: '图一',
      width: 914400, unit: 'emu', coordinate: { pageIndex: null, box: null, unit: null },
      anchor: { offsetX: -91440, relativeFromHorizontal: 'margin', relativeFromVertical: 'paragraph', wrap: 'square', behindText: true } });
    expect(ir.content.floats[1]?.anchor.wrap).toBe('inline');
    expect(ir.content.coverage).toEqual({ breaks: 'observed', sections: 'observed', floats: 'partial' });
    expect(raw.warnings.some((warning) => warning.code === 'UNSUPPORTED_CONTENT')).toBe(true);
    expect(verify({ ir, confidenceFloor: 0.5 }).ok).toBe(true);
  });

  it('honors all three observation flags independently', () => {
    const allOff = parseRawResult(runEngine({ parsePageBreaks: false, parseSections: false, parseFloatingObjects: false }));
    expect(allOff.breaks).toBeUndefined();
    expect(allOff.sections).toBeUndefined();
    expect(allOff.floats).toBeUndefined();
    const content = irFor(allOff).content;
    expect([content.breaks, content.sections, content.floats]).toEqual([[], [], []]);
    const sectionsOff = parseRawResult(runEngine({ parseSections: false }));
    expect(sectionsOff.breaks?.some((item) => item.kind === 'section')).toBe(true);
    const breaksOff = parseRawResult(runEngine({ parsePageBreaks: false }));
    expect(breaksOff.sections).toHaveLength(2);
  });

  it('rejects a floating-object budget before returning a partial success', () => {
    expect(runEngine({}, { maxFloatingObjects: 1 })).toMatchObject({ failure: { reason: 'limit_exceeded', detail: { limit: 'maxFloatingObjects' } } });
    expect(parseRawResult(runEngine({ enforceLimits: false }, { maxFloatingObjects: 1 })).floats).toHaveLength(2);
  });

  it.each(['body', 'kind', 'unit', 'dimension', 'section', 'break'])('rejects invalid observation protocol: %s', (field) => {
    const payload = runEngine();
    if (field === 'body') payload.floats[0].bodyIndex = 99;
    if (field === 'kind') payload.floats[0].kind = 'imaginary';
    if (field === 'unit') payload.floats[0].anchor.unit = 'meters';
    if (field === 'dimension') payload.floats[0].width = -1;
    if (field === 'section') payload.sections[0].startBodyIndex = 99;
    if (field === 'break') payload.breaks[0].beforeBodyIndex = 99;
    expect(() => parseRawResult(payload)).toThrowError(expect.objectContaining({ code: 'ENGINE_PROTOCOL_ERROR' }));
  });

  it.each(['relationship', 'part', 'section', 'break', 'coordinate'])('detects a broken cross-IR reference: %s', (field) => {
    const check = referenceCheck(parseRawResult(runEngine()), (ir) => {
      if (field === 'relationship') ir.content.floats[0]!.relationshipId = 'missing';
      if (field === 'part') ir.parts = ir.parts.filter((part) => !part.name.endsWith('.png'));
      if (field === 'section') ir.content.sections[0]!.startBeforeNodeId = 'missing';
      if (field === 'break') ir.content.breaks[0]!.beforeNodeId = 'missing';
      if (field === 'coordinate') ir.content.floats[0]!.coordinate.part = 'missing';
    });
    expect(check?.status).toBe('fail');
  });
});

describe('real process bridge with deterministic layout double', () => {
  const moduleFor = (mode = 'normal') => createDocxComplexParseModule({ config: {
    engine: { pythonPath: python, env: { PYTHONPATH: resolve('tests/stubs'), DOCX_TEST_LAYOUT: mode } },
  } });
  const input = () => ({ artifactRef: { id: 'fixture', uri: path }, requestId: 'fixture', operation: 'execute' as const });

  it('retains all fragments and indexes a spanning table once per page', async () => {
    const output = await moduleFor().handlers.execute(input());
    const ir = output.result.ir;
    const table = ir.content.blocks[2]!;
    expect(table.fragments).toHaveLength(3);
    expect(table.table).toMatchObject({ spansPages: true, pageRange: [1, 2] });
    expect(table.table?.grid[0]?.[0]?.coordinate?.pageIndex).toBeNull();
    expect(ir.content.pages.map((page) => page.nodeIds)).toEqual([
      [ir.content.blocks[0]!.id], [ir.content.blocks[1]!.id, table.id], [table.id],
    ]);
    expect(verify({ ir, confidenceFloor: 0.5 }).ok).toBe(true);
    ir.content.pages[2]!.nodeIds = [];
    expect(verify({ ir, confidenceFloor: 0.5 }).checks.find((check) => check.id === 'content.references')?.status).toBe('fail');
  });

  it('defaults to structural fallback on layout failure, but honors required geometry', async () => {
    const module = moduleFor('fail');
    const result = await module.handlers.execute(input());
    expect(result.result.ir.content.blocks[0]?.coordinate.pageIndex).toBeNull();
    expect(result.result.ir.content.floats).toHaveLength(2);
    expect(result.warnings.some((warning) => warning.code === 'PAGE_GEOMETRY_UNAVAILABLE')).toBe(true);
    await expect(module.handlers.execute({ ...input(), options: { featureFlags: { requirePageGeometry: true } } }))
      .rejects.toMatchObject({ code: 'LAYOUT_UNAVAILABLE' });
  });

  it('preserves disabled status through handlers and does not attach geometry to drawings', async () => {
    const output = await moduleFor().handlers.execute({ ...input(), options: { featureFlags: {
      parsePageBreaks: false, parseSections: false, parseFloatingObjects: false,
    } } });
    expect(output.result.ir.content.coverage).toEqual({ breaks: 'disabled', sections: 'disabled', floats: 'disabled' });
    expect(output.result.ir.content.floats).toEqual([]);
    const enabled = await moduleFor().handlers.execute(input());
    expect(enabled.result.ir.content.floats.every((item) => item.coordinate.pageIndex === null && item.coordinate.box === null)).toBe(true);
  });

  it('enforces and reports float budgets through the public handlers', async () => {
    const module = moduleFor();
    await expect(module.handlers.execute({ ...input(), options: { limits: { maxFloatingObjects: 1 } } }))
      .rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const output = await module.handlers.execute({ ...input(), options: {
      limits: { maxFloatingObjects: 1 }, featureFlags: { enforceLimits: false },
    } });
    expect(output.warnings.some((warning) => warning.code === 'LIMIT_APPLIED')).toBe(true);
    expect(output.telemetry?.floatingObjectCount).toBe(2);
  });

  it('keeps older v1 engines compatible and validates injected engines too', async () => {
    const legacy = rawResult([paragraph('legacy')]);
    const module = createDocxComplexParseModule({ engine: { name: 'legacy', parse: async () => legacy, dispose: async () => {} } });
    const output = await module.handlers.execute(input());
    expect(output.result.ir.content.floats).toEqual([]);
    expect(output.result.ir.content.coverage?.floats).toBe('unavailable');
    legacy.floats = [{ bodyIndex: 99 } as never];
    await expect(module.handlers.execute(input())).rejects.toMatchObject({ code: 'ENGINE_PROTOCOL_ERROR' });
  });
});

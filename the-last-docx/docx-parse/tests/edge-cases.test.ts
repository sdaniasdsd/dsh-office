/**
 * 边界测试 —— 验证「不合规/恶意/畸形输入」被稳稳接住，且失败原因可编程判断。
 *
 * 这一组测试把「引擎输出」当作普通数据喂给模块，因此既能覆盖真实的 Python 引擎，
 * 也能用测试替身精确构造只在生产中偶发的畸形结果。
 */
import { existsSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { createDocxParseModule } from '../src/index';
import type { ExecuteInput, InspectInput, VerifyInput } from '../src/index';
import { permissivePolicy, artifactRef, FakeEngine, fixture, parseResult } from './support';

/** 捕获 handler 抛出的模块错误码（失败即测试失败）。 */
async function errorCodeOf(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    expect(error).toHaveProperty('code');
    return (error as { code: string }).code;
  }
  throw new Error('expected the call to reject, but it resolved');
}

/** 构造一个使用替身引擎的模块。 */
function moduleWith(overrides: Parameters<typeof parseResult>[0] = {}) {
  return createDocxParseModule({ engine: new FakeEngine(parseResult(overrides)) });
}

function inspectInput(uri = 'fake.docx', options?: InspectInput['options']): InspectInput {
  const input: InspectInput = {
    artifactRef: artifactRef(uri),
    operation: 'inspect',
    requestId: 'edge-1',
  };
  if (options !== undefined) input.options = options;
  return input;
}

describe('edge: result adjudication with a stub engine', () => {
  it('rejects OLE/CFB containers (legacy .doc or encrypted OOXML)', async () => {
    const module = moduleWith({ container: 'ole', documentKind: 'unknown', documentVariant: null });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('UNSUPPORTED_CONTAINER');
  });

  it('rejects RTF streams', async () => {
    const module = moduleWith({ container: 'rtf', documentKind: 'unknown', documentVariant: null });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('UNSUPPORTED_CONTAINER');
  });

  it('rejects unrecognizable containers as a format mismatch', async () => {
    const module = moduleWith({ container: 'unknown', documentKind: 'unknown', documentVariant: null });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('FORMAT_MISMATCH');
  });

  it('rejects a valid OPC package that is not WordprocessingML', async () => {
    const module = moduleWith({ documentKind: 'spreadsheetml', documentVariant: null, blocks: [] });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('FORMAT_MISMATCH');
  });

  it('rejects an empty package', async () => {
    const module = moduleWith({ partCount: 0, parts: [] });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('PARSE_FAILED');
  });

  it('translates engine error kinds into module error codes', async () => {
    const badZip = moduleWith({ error: { kind: 'BAD_ZIP', message: 'corrupt' } });
    expect(await errorCodeOf(badZip.handlers.inspect(inspectInput()))).toBe('PARSE_FAILED');

    const notFound = moduleWith({ error: { kind: 'NOT_FOUND', message: 'gone' } });
    expect(await errorCodeOf(notFound.handlers.inspect(inspectInput()))).toBe('ARTIFACT_NOT_FOUND');
  });

  it('promotes a budget hit to a fatal error when limits are enforced', async () => {
    const module = moduleWith({ limitHit: 'maxEntryUncompressedBytes' });
    expect(await errorCodeOf(module.handlers.inspect(inspectInput()))).toBe('LIMIT_EXCEEDED');
  });

  it('downgrades a budget hit to a warning when limits are not enforced', async () => {
    const module = moduleWith({ limitHit: 'maxEntryUncompressedBytes' });
    const output = await module.handlers.inspect(
      inspectInput('fake.docx', { featureFlags: { enforceLimits: false } }),
    );
    expect(output.warnings.map((warning) => warning.code)).toContain('LIMIT_APPLIED');
  });
});

describe('edge: empty and partial documents', () => {
  it('treats an empty body as a warning, not an error', async () => {
    const module = moduleWith({ blocks: [], issues: [{ code: 'EMPTY_DOCUMENT', message: 'no body' }] });
    const output = await module.handlers.inspect(inspectInput());
    const codes = output.warnings.map((warning) => warning.code);
    expect(codes).toContain('EMPTY_DOCUMENT');
    expect(output.result.format).toBe('docx');
  });

  it('keeps unknown engine issue codes as partially-parsed warnings', async () => {
    const module = moduleWith({ issues: [{ code: 'SOMETHING_NEW', message: 'unknown' }] });
    const output = await module.handlers.inspect(inspectInput());
    const warning = output.warnings.find((entry) => entry.details?.['issueCode'] === 'SOMETHING_NEW');
    expect(warning?.code).toBe('PARTIALLY_PARSED');
  });
});

describe('edge: declared-hint conflicts', () => {
  it('hard-fails when the declared extension is outside the Word family', async () => {
    const module = moduleWith();
    const input = inspectInput('fake.docx', { declaredExtension: 'xlsx' });
    expect(await errorCodeOf(module.handlers.inspect(input))).toBe('FORMAT_MISMATCH');
  });

  it('soft-warns when docx is declared for a macro-enabled document', async () => {
    const module = moduleWith({ documentVariant: 'macroEnabled', extension: 'docm' });
    const output = await module.handlers.inspect(
      inspectInput('fake.docm', { declaredExtension: 'docx' }),
    );
    expect(output.result.format).toBe('docm');
    expect(output.warnings.map((warning) => warning.code)).toContain('DECLARED_HINT_MISMATCH');
  });

  it('treats application/msword as a soft conflict rather than a mismatch', async () => {
    const module = moduleWith();
    const output = await module.handlers.inspect(
      inspectInput('fake.doc', { declaredMimeType: 'application/msword' }),
    );
    expect(output.warnings.map((warning) => warning.code)).toContain('DECLARED_HINT_MISMATCH');
  });
});

describe('edge: real engine against malformed artifacts', () => {
  const module = createDocxParseModule();

  it('reports a missing artifact', async () => {
    const input = inspectInput(fixture('definitely-missing.docx'));
    expect(await errorCodeOf(module.handlers.inspect(input))).toBe('ARTIFACT_NOT_FOUND');
  });

  it('refuses URIs the default resolver cannot reach', async () => {
    const input = inspectInput('https://example.com/a.docx');
    expect(await errorCodeOf(module.handlers.inspect(input))).toBe('UNSUPPORTED_ARTIFACT_URI');
  });

  it('classifies a PNG payload as a format mismatch', async () => {
    expect(existsSync(fixture('plain.bin'))).toBe(true);
    expect(await errorCodeOf(module.handlers.inspect(inspectInput(fixture('plain.bin'))))).toBe(
      'FORMAT_MISMATCH',
    );
  });

  it('classifies an OLE container as unsupported', async () => {
    expect(await errorCodeOf(module.handlers.inspect(inspectInput(fixture('legacy.doc'))))).toBe(
      'UNSUPPORTED_CONTAINER',
    );
  });

  it('classifies an RTF stream as unsupported', async () => {
    expect(await errorCodeOf(module.handlers.inspect(inspectInput(fixture('sample.rtf'))))).toBe(
      'UNSUPPORTED_CONTAINER',
    );
  });

  it('classifies a truncated ZIP as a parse failure', async () => {
    expect(await errorCodeOf(module.handlers.inspect(inspectInput(fixture('truncated.docx'))))).toBe(
      'PARSE_FAILED',
    );
  });

  it('classifies a spreadsheet package as a format mismatch', async () => {
    expect(await errorCodeOf(module.handlers.inspect(inspectInput(fixture('not-word.docx'))))).toBe(
      'FORMAT_MISMATCH',
    );
  });

  it('classifies a package without [Content_Types].xml as a format mismatch', async () => {
    expect(
      await errorCodeOf(module.handlers.inspect(inspectInput(fixture('no-content-types.zip')))),
    ).toBe('FORMAT_MISMATCH');
  });

  it('enforces a caller-supplied entry-size budget', async () => {
    const input = inspectInput(fixture('bomb.docx'), { limits: { maxEntryUncompressedBytes: 1024 } });
    expect(await errorCodeOf(module.handlers.inspect(input))).toBe('LIMIT_EXCEEDED');
  });

  it('accepts a plain document', async () => {
    const output = await module.handlers.inspect(inspectInput(fixture('clean.docx')));
    expect(output.result.format).toBe('docx');
    expect(output.artifacts).toEqual([]);
  });

  it('names each handler operation in the output envelope', async () => {
    const executeInput: ExecuteInput = {
      artifactRef: artifactRef(fixture('clean.docx')),
      operation: 'execute',
      requestId: 'edge-execute',
    };
    const executeOutput = await module.handlers.execute(executeInput);
    expect(executeOutput.operation).toBe('execute');

    const verifyInput: VerifyInput = {
      artifactRef: artifactRef(fixture('clean.docx')),
      operation: 'verify',
      requestId: 'edge-verify',
      policy: permissivePolicy(),
    };
    const verifyOutput = await module.handlers.verify(verifyInput);
    expect(verifyOutput.operation).toBe('verify');
  });
});

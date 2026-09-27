/**
 * 信任边界的规格：**引擎说什么**与**模块信什么**必须分开。
 *
 * 这里的每个断言都对应一类「引擎违约」——它不该被静默接受，也不该被误报成
 * 「文档坏了」。违约一律是 `ENGINE_PROTOCOL_ERROR`。
 */
import { describe, expect, it } from 'vitest';

import {
  parseEngineFailure,
  parseRawResult,
  SUPPORTED_PARSE_VERSION,
} from '../src/domain/docx-complex-parse';
import { DocxComplexParseError } from '../src/errors';

/** 一份最小的合法载荷。 */
function minimal(): Record<string, unknown> {
  return {
    parseVersion: SUPPORTED_PARSE_VERSION,
    mainPart: 'word/document.xml',
    profile: {
      container: 'zip',
      format: 'docx',
      encrypted: false,
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      signatures: ['magic:zip', 'part:word/document.xml'],
    },
    parts: [{ name: 'word/document.xml', sizeBytes: 1024, compressedSize: 512 }],
    relationships: [],
    indicators: { macros: [], externalReferences: [], embeddedObjects: [] },
    blocks: [{ kind: 'paragraph', path: '/w:document/w:body/w:p[1]', text: 'hi' }],
    layout: { available: false, unit: 'pt', pages: [], fragments: [] },
  };
}

describe('parseRawResult：接受', () => {
  it('最小载荷可用，可选字段回落到默认值', () => {
    const result = parseRawResult(minimal());
    expect(result.mainPart).toBe('word/document.xml');
    expect(result.warnings).toEqual([]);
    expect(result.layout.available).toBe(false);
  });

  it('表格块的可选字段有默认值，但 rows 必须显式给出', () => {
    const payload = minimal();
    payload['blocks'] = [{ kind: 'table', path: '/w:document/w:body/w:tbl[1]', rows: [] }];
    const result = parseRawResult(payload);
    const block = result.blocks[0];
    expect(block?.kind).toBe('table');
    if (block?.kind !== 'table') throw new Error('expected table');
    expect(block.tblGridColumns).toBe(0);
    expect(block.headerRowCount).toBe(0);
    expect(block.rows).toEqual([]);
  });

  it('vMerge 缺省时视作「未合并」而非 continue', () => {
    const payload = minimal();
    payload['blocks'] = [
      {
        kind: 'table',
        path: '/w:document/w:body/w:tbl[1]',
        tblGridColumns: 1,
        rows: [{ cells: [{ text: 'a' }] }],
      },
    ];
    const result = parseRawResult(payload);
    const block = result.blocks[0];
    if (block?.kind !== 'table') throw new Error('expected table');
    expect(block.rows[0]?.cells[0]?.vMerge).toBeNull();
    expect(block.rows[0]?.cells[0]?.gridSpan).toBe(1);
  });
});

describe('parseRawResult：拒绝引擎违约', () => {
  const reject = (mutate: (payload: Record<string, unknown>) => void, fragment: string) => {
    const payload = minimal();
    mutate(payload);
    let caught: unknown;
    try {
      parseRawResult(payload);
    } catch (error) {
      caught = error;
    }
    expect(caught, `expected ${fragment} to be rejected`).toBeInstanceOf(DocxComplexParseError);
    expect((caught as DocxComplexParseError).code).toBe(fragment);
  };

  it('协议版本不符', () => {
    reject((payload) => { payload['parseVersion'] = 99; }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('顶层不是对象', () => {
    let caught: unknown;
    try {
      parseRawResult('nope');
    } catch (error) {
      caught = error;
    }
    expect((caught as DocxComplexParseError).code).toBe('ENGINE_PROTOCOL_ERROR');
  });

  it('gridSpan 为 0 或负数', () => {
    reject((payload) => {
      payload['blocks'] = [
        { kind: 'table', path: '/t', rows: [{ cells: [{ text: 'a', gridSpan: 0 }] }] },
      ];
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('vMerge 取值不在枚举内', () => {
    reject((payload) => {
      payload['blocks'] = [
        { kind: 'table', path: '/t', rows: [{ cells: [{ text: 'a', vMerge: 'sideways' }] }] },
      ];
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('表格块缺 rows', () => {
    // 表格不报行，等于引擎没说清结构；与其猜成空表，不如按违约处理。
    reject((payload) => {
      payload['blocks'] = [{ kind: 'table', path: '/t' }];
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('告警码不在本模块固定的表内', () => {
    reject((payload) => {
      payload['warnings'] = [{ code: 'MADE_UP_CODE', message: 'nope' }];
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('版面自称可用却没有片段', () => {
    // 引擎说「我有版面」却不给数据，比诚实地说「没有」更危险：宁可降级为不可用。
    const payload = minimal();
    payload['layout'] = { available: true, unit: 'pt', pages: [], fragments: [] };
    expect(parseRawResult(payload).layout.available).toBe(false);
  });

  it('版面坐标不是有限数', () => {
    reject((payload) => {
      payload['layout'] = {
        available: true,
        unit: 'pt',
        pages: [{ index: 0, width: 612, height: 792 }],
        fragments: [{ bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: Number.NaN, y: 0, width: 1, height: 1 }],
      };
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('缺包级事实', () => {
    // 包级事实缺失意味着引擎没有真正读过容器——此时任何「无宏 / 无外部引用」
    // 的安全结论都无从谈起，只能按违约处理。
    reject((payload) => { delete payload['profile']; }, 'ENGINE_PROTOCOL_ERROR');
    reject((payload) => { delete payload['parts']; }, 'ENGINE_PROTOCOL_ERROR');
    reject((payload) => { delete payload['relationships']; }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('indicators 少一类安全指标', () => {
    reject((payload) => {
      payload['indicators'] = { macros: [], externalReferences: [] };
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('格式身份不在枚举内', () => {
    reject((payload) => {
      payload['profile'] = {
        container: 'zip',
        format: 'pdf',
        encrypted: false,
        mediaType: null,
        signatures: [],
      };
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('encrypted 不是布尔也不是 null', () => {
    reject((payload) => {
      payload['profile'] = {
        container: 'zip',
        format: 'docx',
        encrypted: 'maybe',
        mediaType: null,
        signatures: [],
      };
    }, 'ENGINE_PROTOCOL_ERROR');
  });

  it('宏指标的 kind 不在枚举内', () => {
    reject((payload) => {
      payload['indicators'] = {
        macros: [{ kind: 'dde', part: 'word/vbaProject.bin', sizeBytes: 1 }],
        externalReferences: [],
        embeddedObjects: [],
      };
    }, 'ENGINE_PROTOCOL_ERROR');
  });
});

describe('parseRawResult：包级事实的边界取值', () => {
  it('encrypted=null 表示「无法判断」，不被折算成 false', () => {
    const payload = minimal();
    payload['profile'] = {
      container: 'zip',
      format: 'docx',
      encrypted: null,
      mediaType: null,
      signatures: ['magic:zip'],
    };
    // 折算成 false 会让策略层以为「已确认未加密」，这是安全语义上的降级。
    expect(parseRawResult(payload).profile.encrypted).toBeNull();
  });

  it('mediaType 缺失时记为 null 而不是空串', () => {
    const payload = minimal();
    payload['profile'] = {
      container: 'zip',
      format: 'unknown',
      encrypted: false,
      signatures: [],
    };
    const profile = parseRawResult(payload).profile;
    expect(profile.mediaType).toBeNull();
    expect(profile.signatures).toEqual([]);
  });

  it('外部引用的 targetMode 缺省为 Internal', () => {
    const payload = minimal();
    payload['relationships'] = [{ id: 'rId1', type: 't', target: 'word/x.xml' }];
    expect(parseRawResult(payload).relationships[0]?.targetMode).toBe('Internal');
  });
});

describe('parseEngineFailure', () => {
  it('成功载荷不是失败信封', () => {
    expect(parseEngineFailure(minimal())).toBeNull();
    expect(parseEngineFailure('nope')).toBeNull();
    expect(parseEngineFailure([])).toBeNull();
  });

  it('解析出原因与上下文', () => {
    const failure = parseEngineFailure({
      failure: { reason: 'limit_exceeded', message: 'too big', detail: { limit: 'maxBlocks' } },
    });
    expect(failure?.reason).toBe('limit_exceeded');
    expect(failure?.detail).toEqual({ limit: 'maxBlocks' });
  });

  it('信封残缺时抛协议错误', () => {
    expect(() => parseEngineFailure({ failure: { reason: 'x' } })).toThrow(DocxComplexParseError);
  });
});

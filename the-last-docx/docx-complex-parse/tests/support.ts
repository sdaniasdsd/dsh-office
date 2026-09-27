/**
 * 测试构造器：把「引擎会交上来的原始观察」写成可读的短表达式。
 *
 * 这些 helper 刻意与 Python 侧的输出同形，因此测试不需要起进程、不需要 fixture
 * 文件，就能覆盖网格展开的全部分支。
 */
import type {
  RawBlock,
  RawLayout,
  RawParseResult,
  RawTableCell,
  RawTableRow,
} from '../src/domain/docx-complex-parse';

/** 一个原始格。`vMerge` 三态：null=不合并，restart=纵向起点，continue=纵向续接。 */
export function cell(
  text: string,
  gridSpan = 1,
  vMerge: 'restart' | 'continue' | null = null,
): RawTableCell {
  return { gridSpan, vMerge, text };
}

/** 一个原始行。 */
export function row(
  cells: RawTableCell[],
  gridBefore = 0,
  gridAfter = 0,
): RawTableRow {
  return { gridBefore, gridAfter, cells };
}

/** 一个原始表格块。 */
export function table(
  rows: RawTableRow[],
  tblGridColumns: number,
  headerRowCount = 0,
): RawBlock {
  return {
    kind: 'table',
    path: '/w:document/w:body/w:tbl[1]',
    tblGridColumns,
    headerRowCount,
    rows,
  };
}

/** 一个原始段落块。 */
export function paragraph(text: string, ordinal = 1): RawBlock {
  return { kind: 'paragraph', path: `/w:document/w:body/w:p[${ordinal}]`, text };
}

/** 一份完整的原始观察。版面默认不可用，需要时显式传入。 */
export function rawResult(
  blocks: RawBlock[],
  options: {
    layout?: Partial<RawLayout>;
    warnings?: RawParseResult['warnings'];
    mainPart?: string;
    profile?: Partial<RawParseResult['profile']>;
    parts?: RawParseResult['parts'];
    relationships?: RawParseResult['relationships'];
    indicators?: Partial<RawParseResult['indicators']>;
  } = {},
): RawParseResult {
  return {
    parseVersion: 1,
    mainPart: options.mainPart ?? 'word/document.xml',
    profile: {
      container: 'zip',
      format: 'docx',
      encrypted: false,
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      // 默认给全三个信号，这样 `toFormatProfile` 的置信度是满分，
      // 需要测「信号不足」的用例再显式删减。
      signatures: ['magic:zip', 'part:word/document.xml', 'content-types:wordprocessingml'],
      ...options.profile,
    },
    parts: options.parts ?? [
      { name: 'word/document.xml', sizeBytes: 1024, compressedSize: 512 },
    ],
    relationships: options.relationships ?? [],
    indicators: {
      macros: [],
      externalReferences: [],
      embeddedObjects: [],
      ...options.indicators,
    },
    blocks,
    warnings: options.warnings ?? [],
    layout: {
      available: false,
      unit: 'pt',
      pages: [],
      fragments: [],
      ...options.layout,
    },
  };
}

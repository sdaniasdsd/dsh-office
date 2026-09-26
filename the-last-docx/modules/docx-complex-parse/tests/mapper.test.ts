/**
 * 归一化的规格：引擎的形态到此为止，出去的一律是 contract.ts 定义的类型。
 */
import { describe, expect, it } from 'vitest';

import { COMPLEX_LAYOUT_SCHEME, type ComplexTable } from '../src/contract';
import { toComplexContent } from '../src/mapper';
import { cell, paragraph, rawResult, row, table } from './support';

function tableOf(index: number, blocks: ReturnType<typeof toComplexContent>['content']['blocks']): ComplexTable {
  const found = blocks[index];
  if (!found?.table) throw new Error(`block ${index} is not a table`);
  return found.table;
}

const SQUARE = table([row([cell('a'), cell('b')]), row([cell('c'), cell('d')])], 2);

describe('toComplexContent：身份与坐标', () => {
  it('id 由部件名 + 结构路径合成，稳定且可读', () => {
    const { content } = toComplexContent(rawResult([paragraph('hello')]));
    expect(content.blocks[0]?.id).toBe('word/document.xml!/w:document/w:body/w:p[1]');
    expect(content.mainPart).toBe('word/document.xml');
    expect(content.scheme).toBe(COMPLEX_LAYOUT_SCHEME);
  });

  it('无版面时坐标退化为纯结构路径，而不是编造页号', () => {
    const { content, warnings } = toComplexContent(rawResult([SQUARE]));
    const coordinate = content.blocks[0]?.coordinate;
    expect(coordinate?.pageIndex).toBeNull();
    expect(coordinate?.box).toBeNull();
    expect(coordinate?.unit).toBeNull();
    expect(content.pages).toEqual([]);
    expect(warnings.map((warning) => warning.code)).toContain('PAGE_GEOMETRY_UNAVAILABLE');
  });

  it('有版面时把页号与页内矩形挂到块上', () => {
    const { content } = toComplexContent(
      rawResult([SQUARE], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [{ index: 0, width: 612, height: 792 }],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 100, width: 468, height: 20 },
          ],
        },
      }),
    );
    const coordinate = content.blocks[0]?.coordinate;
    expect(coordinate?.pageIndex).toBe(0);
    expect(coordinate?.box).toEqual({ x: 72, y: 100, width: 468, height: 20 });
    expect(coordinate?.unit).toBe('pt');
  });

  it('单元格坐标由表格路径 + 原始行列下标合成，且不冒充整表的矩形', () => {
    const { content } = toComplexContent(
      rawResult([SQUARE], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [{ index: 0, width: 612, height: 792 }],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 72, width: 468, height: 20 },
          ],
        },
      }),
    );
    const cellCoordinate = tableOf(0, content.blocks).grid[1]?.[0]?.coordinate;
    expect(cellCoordinate?.structuralPath).toBe(
      '/w:document/w:body/w:tbl[1]/w:tr[2]/w:tc[1]',
    );
    expect(cellCoordinate?.pageIndex).toBe(0);
    // 表级矩形不属于单元格——填进去会比「不知道」更糟。
    expect(cellCoordinate?.box).toBeNull();
    // 被覆盖的格子没有坐标。
    const merged = toComplexContent(
      rawResult([table([row([cell('w', 2), cell('x')])], 3)]),
    ).content.blocks[0]?.table;
    expect(merged?.grid[0]?.[1]?.coordinate).toBeNull();
  });
});

describe('toComplexContent：页几何', () => {
  it('由块的外接范围反推页边距，并判方向', () => {
    const { content } = toComplexContent(
      rawResult([paragraph('a'), paragraph('b', 2)], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [{ index: 0, width: 612, height: 792 }],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 72, width: 468, height: 12 },
            { bodyIndex: 1, physicalPage: 1, displayedPage: 1, x: 72, y: 100, width: 468, height: 20 },
          ],
        },
      }),
    );

    expect(content.pages).toHaveLength(1);
    const page = content.pages[0];
    expect(page?.geometry).toEqual({
      width: 612,
      height: 792,
      unit: 'pt',
      orientation: 'portrait',
      margins: { top: 72, right: 72, bottom: 672, left: 72 },
    });
    // 版面来自计算而非文档声明，因此必须标为推断。
    expect(page?.inferred).toBe(true);
    expect(page?.nodeIds).toEqual([
      'word/document.xml!/w:document/w:body/w:p[1]',
      'word/document.xml!/w:document/w:body/w:p[2]',
    ]);
  });

  it('宽大于高时判为横向', () => {
    const { content } = toComplexContent(
      rawResult([paragraph('a')], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [{ index: 0, width: 792, height: 612 }],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 72, width: 648, height: 12 },
          ],
        },
      }),
    );
    expect(content.pages[0]?.geometry?.orientation).toBe('landscape');
  });

  it('同一块有多个片段时判为跨页，并给出页区间', () => {
    const { content } = toComplexContent(
      rawResult([SQUARE], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [
            { index: 0, width: 612, height: 792 },
            { index: 1, width: 612, height: 792 },
          ],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 700, width: 468, height: 40 },
            { bodyIndex: 0, physicalPage: 2, displayedPage: 2, x: 72, y: 72, width: 468, height: 60 },
          ],
        },
      }),
    );
    const complex = tableOf(0, content.blocks);
    expect(complex.spansPages).toBe(true);
    expect(complex.pageRange).toEqual([0, 1]);
    // 块的坐标取首个片段，而非任意一个。
    expect(content.blocks[0]?.coordinate.pageIndex).toBe(0);
    expect(content.pages).toHaveLength(2);
  });

  it('单页的表不会被误报为跨页', () => {
    const { content } = toComplexContent(
      rawResult([SQUARE], {
        layout: {
          available: true,
          unit: 'pt',
          pages: [{ index: 0, width: 612, height: 792 }],
          fragments: [
            { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 72, width: 468, height: 20 },
          ],
        },
      }),
    );
    const complex = tableOf(0, content.blocks);
    expect(complex.spansPages).toBe(false);
    expect(complex.pageRange).toEqual([0, 0]);
  });
});

describe('toComplexContent：置信度与告警', () => {
  it('引擎告警排在 mapper 判断之前，且都保留来源', () => {
    const { warnings } = toComplexContent(
      rawResult([paragraph('a')], {
        warnings: [{ code: 'NESTED_TABLE_FLATTENED', message: 'nested', path: '/t' }],
        layout: { available: false, unit: 'pt', pages: [], fragments: [] },
      }),
    );
    expect(warnings[0]?.code).toBe('NESTED_TABLE_FLATTENED');
    expect(warnings[0]?.path).toBe('/t');
    expect(warnings[1]?.code).toBe('PAGE_GEOMETRY_UNAVAILABLE');
  });

  it('没有任何可见块时给出 EMPTY_DOCUMENT', () => {
    const { warnings, content } = toComplexContent(rawResult([]));
    expect(warnings.map((warning) => warning.code)).toEqual(['EMPTY_DOCUMENT']);
    expect(content.blocks).toEqual([]);
  });

  it('格子由 tblGrid 直接给出时置信度来源为 native', () => {
    const { content } = toComplexContent(rawResult([SQUARE]));
    const complex = tableOf(0, content.blocks);
    expect(complex.confidence.source).toBe('native');
    expect(complex.grid[0]?.[0]?.confidence.source).toBe('native');
    expect(complex.grid[0]?.[0]?.confidence.score).toBeGreaterThan(0.9);
  });

  it('列数靠推导、或文档自相矛盾时降级为 structural', () => {
    const derived = toComplexContent(
      rawResult([table([row([cell('a'), cell('b')])], 0)]),
    ).content.blocks[0]?.table;
    expect(derived?.confidence.source).toBe('structural');

    const inconsistent = toComplexContent(
      rawResult([table([row([cell('a'), cell('b'), cell('c')])], 2)]),
    ).content.blocks[0]?.table;
    expect(inconsistent?.confidence.source).toBe('structural');
    expect(inconsistent?.confidence.notes.join(' ')).toContain('exceeds tblGrid');
  });

  it('续接却找不到起点的格子标为 heuristic', () => {
    const { content } = toComplexContent(
      rawResult([table([row([cell('orphan', 1, 'continue')]), row([cell('b')])], 1)]),
    );
    expect(tableOf(0, content.blocks).grid[0]?.[0]?.confidence.source).toBe('heuristic');
  });
});

describe('toComplexContent：纯函数性', () => {
  it('同一输入产出逐字段相同的结果', () => {
    const input = rawResult([paragraph('a'), SQUARE], {
      layout: {
        available: true,
        unit: 'pt',
        pages: [{ index: 0, width: 612, height: 792 }],
        fragments: [
          { bodyIndex: 0, physicalPage: 1, displayedPage: 1, x: 72, y: 72, width: 468, height: 12 },
          { bodyIndex: 1, physicalPage: 1, displayedPage: 1, x: 72, y: 90, width: 468, height: 40 },
        ],
      },
    });
    expect(JSON.stringify(toComplexContent(input))).toBe(JSON.stringify(toComplexContent(input)));
  });
});

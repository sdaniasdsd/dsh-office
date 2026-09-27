/**
 * 逻辑网格展开的行为规格。
 *
 * 判据来自 WordprocessingML 自身：网格是先建立的，合并只是「相对指令」，
 * 因此每个断言都在问同一件事——**这一格最终归属哪个真实单元格**。
 */
import { describe, expect, it } from 'vitest';

import type { RawTableBlock } from '../src/domain/docx-complex-parse';
import { expandLogicalGrid } from '../src/mapper';
import { cell, row, table } from './support';

function asTable(block: ReturnType<typeof table>): RawTableBlock {
  if (block.kind !== 'table') throw new Error('expected a table block');
  return block;
}

/** 把网格渲染成便于断言的文本矩阵：`·` 表示被合并覆盖。 */
function render(block: RawTableBlock, columnCount: number, rowCount: number): string[][] {
  const grid = expandLogicalGrid(block);
  const out: string[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const line: string[] = [];
    for (let c = 0; c < columnCount; c += 1) {
      const found = grid.grid[r]?.[c];
      line.push(found && found.origin ? found.text : '·');
    }
    out.push(line);
  }
  return out;
}

describe('expandLogicalGrid：无合并', () => {
  it('把规整的 N×M 表原样还原', () => {
    const grid = expandLogicalGrid(
      asTable(table([row([cell('a'), cell('b')]), row([cell('c'), cell('d')])], 2)),
    );

    expect(grid.rowCount).toBe(2);
    expect(grid.columnCount).toBe(2);
    expect(grid.merges).toEqual([]);
    expect(grid.issues).toEqual([]);
    expect(render(asTable(table([row([cell('a'), cell('b')]), row([cell('c'), cell('d')])], 2)), 2, 2))
      .toEqual([['a', 'b'], ['c', 'd']]);
  });

  it('每个真实格子的 origin 指向自身，且带原始下标', () => {
    const grid = expandLogicalGrid(asTable(table([row([cell('a'), cell('b')])], 2)));
    expect(grid.grid[0]?.[0]?.origin).toEqual({ row: 0, column: 0 });
    expect(grid.grid[0]?.[1]?.origin).toEqual({ row: 0, column: 1 });
    expect(grid.grid[0]?.[0]?.sourceIndex).toBe(0);
    expect(grid.grid[0]?.[1]?.sourceIndex).toBe(1);
    expect(grid.grid[0]?.[0]?.coverage).toBe('self');
  });
});

describe('expandLogicalGrid：横向合并', () => {
  it('gridSpan=2 让右邻格变成被覆盖的占位格', () => {
    const block = asTable(table([row([cell('wide', 2), cell('x')]), row([cell('p'), cell('q')])], 3));
    const grid = expandLogicalGrid(block);

    expect(grid.columnCount).toBe(3);
    expect(render(block, 3, 2)).toEqual([['wide', '·', 'x'], ['p', 'q', '·']]);
    expect(grid.grid[0]?.[1]?.origin).toBeNull();
    expect(grid.grid[0]?.[1]?.text).toBe('');
    expect(grid.grid[0]?.[1]?.sourceIndex).toBeNull();
    expect(grid.merges).toEqual([
      { startRow: 0, startColumn: 0, rowSpan: 1, columnSpan: 2, kind: 'gridSpan', text: 'wide' },
    ]);
  });
});

describe('expandLogicalGrid：纵向合并', () => {
  it('restart + continue 合成一块，续接格归属上方起点', () => {
    const block = asTable(
      table(
        [row([cell('top', 1, 'restart'), cell('a')]), row([cell('', 1, 'continue'), cell('b')])],
        2,
      ),
    );
    const grid = expandLogicalGrid(block);

    expect(render(block, 2, 2)).toEqual([['top', 'a'], ['·', 'b']]);
    // 续接格自己也声明了文本时，内容以起点为准——合并区只有一份内容。
    expect(grid.merges).toEqual([
      { startRow: 0, startColumn: 0, rowSpan: 2, columnSpan: 1, kind: 'vMerge', text: 'top' },
    ]);
  });

  it('续接链可以跨 3 行以上（取锚点而非上一格，链条才不会断）', () => {
    const block = asTable(
      table(
        [
          row([cell('head', 1, 'restart')]),
          row([cell('', 1, 'continue')]),
          row([cell('', 1, 'continue')]),
        ],
        1,
      ),
    );
    const grid = expandLogicalGrid(block);
    expect(render(block, 1, 3)).toEqual([['head'], ['·'], ['·']]);
    expect(grid.merges[0]?.rowSpan).toBe(3);
  });

  it('找不到起点的 continue 降级为独立格，并记录 MERGE_INCONSISTENT', () => {
    const block = asTable(table([row([cell('orphan', 1, 'continue')]), row([cell('b')])], 1));
    const grid = expandLogicalGrid(block);

    // 宁可多留一格，也不把文本无声吞掉。
    expect(render(block, 1, 2)).toEqual([['orphan'], ['b']]);
    expect(grid.grid[0]?.[0]?.coverage).toBe('orphan');
    expect(grid.issues).toHaveLength(1);
    expect(grid.issues[0]?.code).toBe('MERGE_INCONSISTENT');
  });
});

describe('expandLogicalGrid：横纵同时合并', () => {
  it('2×2 区域标记为 both，并给出完整矩形', () => {
    const block = asTable(
      table(
        [
          row([cell('b00'), cell('b01'), cell('b02')]),
          row([cell('b10'), cell('b11', 2, 'restart')]),
          row([cell('b20'), cell('', 2, 'continue')]),
          row([cell('b30'), cell('b31'), cell('b32')]),
        ],
        3,
      ),
    );
    const grid = expandLogicalGrid(block);

    expect(render(block, 3, 4)).toEqual([
      ['b00', 'b01', 'b02'],
      ['b10', 'b11', '·'],
      ['b20', '·', '·'],
      ['b30', 'b31', 'b32'],
    ]);
    expect(grid.merges).toEqual([
      { startRow: 1, startColumn: 1, rowSpan: 2, columnSpan: 2, kind: 'both', text: 'b11' },
    ]);
  });
});

describe('expandLogicalGrid：行首/行尾偏移', () => {
  it('gridBefore 让整行右移，而不是左对齐补齐', () => {
    const block = asTable(
      table(
        [
          row([cell('s00'), cell('s01'), cell('s02')]),
          row([cell('s11'), cell('s12')], 1),
          row([cell('s20'), cell('s21'), cell('s22')]),
        ],
        3,
      ),
    );
    const grid = expandLogicalGrid(block);

    expect(render(block, 3, 3)).toEqual([
      ['s00', 's01', 's02'],
      ['·', 's11', 's12'],
      ['s20', 's21', 's22'],
    ]);
    expect(grid.grid[1]?.[1]?.sourceIndex).toBe(0);
    expect(grid.issues).toEqual([]);
  });

  it('gridAfter 只影响行尾占位，不影响后续行的列数', () => {
    const block = asTable(
      table(
        [
          row([cell('t00'), cell('t01'), cell('t02')]),
          row([cell('t10'), cell('t11')], 0, 1),
        ],
        3,
      ),
    );
    const grid = expandLogicalGrid(block);
    expect(grid.columnCount).toBe(3);
    expect(render(block, 3, 2)).toEqual([['t00', 't01', 't02'], ['t10', 't11', '·']]);
  });
});

describe('expandLogicalGrid：tblGrid 与实际情况不符', () => {
  it('tblGrid 缺失时按最长行推导列数', () => {
    const grid = expandLogicalGrid(
      asTable(table([row([cell('a'), cell('b')]), row([cell('c'), cell('d'), cell('e')])], 0)),
    );
    expect(grid.columnCount).toBe(3);
    expect(grid.issues).toEqual([]);
  });

  it('行比 tblGrid 更宽时扩列而不是丢格，并记录不一致', () => {
    const grid = expandLogicalGrid(
      asTable(table([row([cell('a'), cell('b'), cell('c')])], 2)),
    );
    expect(grid.columnCount).toBe(3);
    expect(grid.grid[0]?.[2]?.origin).toEqual({ row: 0, column: 2 });
    expect(grid.issues).toHaveLength(1);
    expect(grid.issues[0]?.message).toContain('exceeds tblGrid 2');
  });

  it('行比 tblGrid 更窄时保留声明列数（占位格留空）', () => {
    const grid = expandLogicalGrid(
      asTable(table([row([cell('a')]), row([cell('b'), cell('c'), cell('d')])], 3)),
    );
    expect(grid.columnCount).toBe(3);
    expect(grid.grid[0]?.[1]?.origin).toBeNull();
    expect(grid.grid[0]?.[2]?.origin).toBeNull();
  });

  it('纵向合并中途改宽度会被记录为不一致', () => {
    const grid = expandLogicalGrid(
      asTable(
        table(
          [
            row([cell('a', 2, 'restart')]),
            row([cell('b', 1, 'continue')]),
          ],
          2,
        ),
      ),
    );
    expect(grid.issues.some((issue) => issue.message.includes('changes width'))).toBe(true);
  });
});

describe('expandLogicalGrid：维度不变式', () => {
  const blocks: RawTableBlock[] = [
    asTable(table([], 0)),
    asTable(table([row([])], 0)),
    asTable(table([row([cell('only')])], 1)),
    asTable(table([row([cell('x', 3)])], 3)),
    asTable(table([row([cell('x')], 2)], 3)),
  ];

  it('网格恒为 rowCount × columnCount 的矩形', () => {
    for (const block of blocks) {
      const grid = expandLogicalGrid(block);
      expect(grid.grid).toHaveLength(grid.rowCount);
      for (const line of grid.grid) {
        expect(line).toHaveLength(grid.columnCount);
      }
      for (const merge of grid.merges) {
        expect(merge.rowSpan).toBeGreaterThanOrEqual(1);
        expect(merge.columnSpan).toBeGreaterThanOrEqual(1);
        expect(merge.startRow + merge.rowSpan).toBeLessThanOrEqual(grid.rowCount);
        expect(merge.startColumn + merge.columnSpan).toBeLessThanOrEqual(grid.columnCount);
      }
    }
  });
});

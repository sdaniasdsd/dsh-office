/**
 * mapper —— 底层原始观察 → 本模块 IR 的【唯一】归一化点。
 *
 * spec 要求「复杂表格、跨页关系、浮动对象、来源坐标、置信度」都要经 mapper
 * 归一化，本文件是这条要求的落点。两条纪律：
 *
 *   1. 本文件是纯函数集合：不读文件、不起进程、不读时间，因此可被反复调用并
 *      在离线测试里逐字节比对；
 *   2. 引擎的形态（原始格、片段）到此为止——出去的只有 contract.ts 定义的类型，
 *      不会再有第二个地方需要理解「引擎说了什么」。
 */
import {
  COMPLEX_LAYOUT_SCHEME,
  type ComplexBlock,
  type ComplexContent,
  type ComplexNodeId,
  type ComplexTable,
  type Confidence,
  type ConfidenceSource,
  type CoordinateUnit,
  type DocxComplexParseIR,
  type DocxComplexParseWarningCode,
  type GridCell,
  type LimitConfig,
  type MergeKind,
  type MergeRegion,
  type ModuleConfig,
  type PageInfo,
  type SourceCoordinate,
  type FeatureFlags,
  type PageBreak,
  type FloatingObject,
} from './contract';
import type { ArtifactRef, FormatProfile, Warning, WarningSeverity } from 'office-core';
import { DocxComplexParseError } from './errors';
import type {
  RawBlock,
  RawFormatKind,
  RawLayout,
  RawLayoutFragment,
  RawLayoutPage,
  RawParseResult,
  RawTableBlock,
  RawTableCell,
} from './domain/docx-complex-parse';

/* -------------------------------------------------------------------------- */
/* 逻辑网格展开（参照 python-docx 的 layout grid 模型）                          */
/* -------------------------------------------------------------------------- */

/** 一处网格异常。它不是错误，而是「文档自相矛盾，已按最保守方式解释」的记录。 */
export interface GridIssue {
  code: 'MERGE_INCONSISTENT';
  message: string;
}

/** 网格中的一个格子。 */
export interface LogicalCell {
  row: number;
  column: number;
  /** 单元格文本；被合并覆盖的格子为空字符串。 */
  text: string;
  /**
   * 承载该格的真实单元格坐标（=自身）。被左侧或上方合并区域覆盖时为 null。
   *
   * 之所以是 null 而不是「指向覆盖它的那个格」：`origin !== null` 因此可以直接
   * 当作「这一格是真实单元格」的判据，调用方无需再比对坐标。要找回覆盖者，
   * 用 `merges`——它给出了区域矩形。
   */
  origin: { row: number; column: number } | null;
  /** 真实单元格在原始行 `cells` 里的下标；被覆盖时为 null。 */
  sourceIndex: number | null;
  /** 归属判定的依据，决定该格的置信度来源。 */
  coverage: 'self' | 'merged' | 'orphan';
}

/** 一处已还原为逻辑网格坐标的合并区域。 */
export interface LogicalMerge {
  startRow: number;
  startColumn: number;
  rowSpan: number;
  columnSpan: number;
  kind: MergeKind;
  text: string;
}

/** 展开后的逻辑网格。 */
export interface LogicalGrid {
  rowCount: number;
  columnCount: number;
  grid: LogicalCell[][];
  merges: LogicalMerge[];
  issues: GridIssue[];
}

/** 一个原始格在行内的落位。 */
interface Placement {
  column: number;
  span: number;
  cellIndex: number;
  cell: RawTableCell;
}

/** 在某一行的落位表里，找出「起始列恰好为 column」的那个格。 */
function anchorAt(entries: readonly Placement[], column: number): Placement | null {
  for (const entry of entries) {
    if (entry.column === column) return entry;
  }
  return null;
}

/**
 * 把 `w:tblGrid` + `w:gridSpan` + `w:vMerge` 展开成矩形逻辑网格。
 *
 * 模型（与 WordprocessingML 一致）：**先有一张网格，格子落在网格上**。
 * 因此这不是「从合并属性反推网格」，而是三步：
 *
 *   1. 定列数——`tblGrid` 声明了它；缺失时按最长行推导；
 *   2. 横向落位——每行用游标走一遍，`gridSpan` 占多列，`gridBefore` 让
 *      行首可以不在第 0 列（这是主流实现常漏的一项）；
 *   3. 纵向续接——`vMerge="continue"` 归属上方同列的锚点，链条可跨多行。
 *
 * 全程不做启发式猜测：所有决定都来自文档里写着的事实，唯一的例外是
 * 「`vMerge="continue"` 却找不到起点」这种自相矛盾的情形，此时按独立格处理
 * 并记录 `MERGE_INCONSISTENT`，宁可多留一格，也不把内容无声吞掉。
 */
export function expandLogicalGrid(table: RawTableBlock): LogicalGrid {
  const issues: GridIssue[] = [];
  const rowCount = table.rows.length;

  // 第一步：把每行原始格换算成「起始列 + 跨度」，并记录观测到的最宽行。
  const placements: Placement[][] = [];
  let observedWidth = 0;

  for (let r = 0; r < rowCount; r += 1) {
    const row = table.rows[r];
    const entries: Placement[] = [];
    let cursor = row ? row.gridBefore : 0;

    if (row) {
      for (let i = 0; i < row.cells.length; i += 1) {
        const cell = row.cells[i];
        if (!cell) continue;
        const span = cell.gridSpan >= 1 ? cell.gridSpan : 1;
        entries.push({ column: cursor, span, cellIndex: i, cell });
        cursor += span;
      }
      cursor += row.gridAfter;
    }

    if (cursor > observedWidth) observedWidth = cursor;
    placements.push(entries);
  }

  // 第二步：定列数。`tblGrid` 权威，但绝不因它而丢格——观测更宽时以观测为准。
  const declared = table.tblGridColumns > 0 ? table.tblGridColumns : 0;
  const columnCount = Math.max(declared, observedWidth);

  if (declared > 0 && observedWidth > declared) {
    issues.push({
      code: 'MERGE_INCONSISTENT',
      message: `observed row width ${observedWidth} exceeds tblGrid ${declared}; grid expanded to ${columnCount} columns`,
    });
  }

  // 第三步：逐格求归属。anchorGrid 记的是「该格最终归属的锚点」。
  // orphanKeys 记下「声明续接却找不到起点」的格子，它们的归属是猜的。
  const orphanKeys = new Set<string>();
  const anchorGrid: ({ row: number; column: number } | null)[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const line: ({ row: number; column: number } | null)[] = [];
    for (let c = 0; c < columnCount; c += 1) line.push(null);
    anchorGrid.push(line);
  }

  for (let r = 0; r < rowCount; r += 1) {
    for (const entry of placements[r] ?? []) {
      const continuing = entry.cell.vMerge === 'continue';
      for (let k = 0; k < entry.span; k += 1) {
        const col = entry.column + k;
        if (col >= columnCount) continue;
        const line = anchorGrid[r];
        if (!line) continue;

        if (!continuing) {
          // 起点格（含 vMerge="restart"）：整段横向跨度都归它。
          line[col] = { row: r, column: entry.column };
          continue;
        }

        // 续接格：取「上方同列」的锚点。取锚点而非上格自身，续接链才能跨多行。
        const above = r > 0 ? anchorGrid[r - 1]?.[col] ?? null : null;
        if (above) {
          line[col] = above;
        } else {
          line[col] = { row: r, column: entry.column };
          orphanKeys.add(`${r}:${col}`);
          issues.push({
            code: 'MERGE_INCONSISTENT',
            message: `vMerge="continue" at row ${r}, column ${col} has no restart above; treated as an independent cell`,
          });
        }
      }
    }
  }

  // 第四步：按归属关系生成格子。
  const grid: LogicalCell[][] = [];
  for (let r = 0; r < rowCount; r += 1) {
    const line: LogicalCell[] = [];
    const entries = placements[r] ?? [];
    for (let c = 0; c < columnCount; c += 1) {
      const anchor = anchorGrid[r]?.[c] ?? null;
      const isOrigin = anchor !== null && anchor.row === r && anchor.column === c;

      if (!isOrigin) {
        line.push({
          row: r,
          column: c,
          text: '',
          origin: null,
          sourceIndex: null,
          coverage: 'merged',
        });
        continue;
      }

      const source = anchorAt(entries, c);
      line.push({
        row: r,
        column: c,
        text: source ? source.cell.text : '',
        origin: { row: r, column: c },
        sourceIndex: source ? source.cellIndex : null,
        coverage: orphanKeys.has(`${r}:${c}`) ? 'orphan' : 'self',
      });
    }
    grid.push(line);
  }

  // 第五步：还原合并区域。只从「起点格」出发数跨度，因此不会重复计数。
  const merges: LogicalMerge[] = [];
  for (let r = 0; r < rowCount; r += 1) {
    for (const entry of placements[r] ?? []) {
      if (entry.cell.vMerge === 'continue') continue;

      let rowSpan = 1;
      for (let rr = r + 1; rr < rowCount; rr += 1) {
        const below = anchorAt(placements[rr] ?? [], entry.column);
        if (!below || below.cell.vMerge !== 'continue') break;
        if (below.span !== entry.span) {
          issues.push({
            code: 'MERGE_INCONSISTENT',
            message: `vertical merge at row ${r}, column ${entry.column} changes width from ${entry.span} to ${below.span} on row ${rr}`,
          });
        }
        rowSpan += 1;
      }

      if (entry.span === 1 && rowSpan === 1) continue;

      merges.push({
        startRow: r,
        startColumn: entry.column,
        rowSpan,
        columnSpan: entry.span,
        kind: entry.span > 1 && rowSpan > 1 ? 'both' : entry.span > 1 ? 'gridSpan' : 'vMerge',
        text: entry.cell.text,
      });
    }
  }

  return { rowCount, columnCount, grid, merges, issues };
}

/* -------------------------------------------------------------------------- */
/* 版面索引                                                                    */
/* -------------------------------------------------------------------------- */

/** 一页上全部正文片段的外接范围，用于反推页边距。 */
interface PageExtent {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

/** 版面索引：把「第几个正文块」映射到「在哪一页、页面哪个矩形」。 */
export interface LayoutIndex {
  available: boolean;
  unit: CoordinateUnit;
  pageByIndex: Map<number, RawLayoutPage>;
  /** 一个正文块的全部片段——跨页的块会有多个，这是判定 `spansPages` 的依据。 */
  fragmentsByBodyIndex: Map<number, RawLayoutFragment[]>;
  extentByPage: Map<number, PageExtent>;
}

/** 取一个块的首个片段（页序号最小、同页 y 最小），代表块的起始位置。 */
function primaryFragment(fragments: readonly RawLayoutFragment[] | undefined): RawLayoutFragment | null {
  if (!fragments || fragments.length === 0) return null;
  let best: RawLayoutFragment | null = null;
  for (const fragment of fragments) {
    if (
      best === null ||
      fragment.physicalPage < best.physicalPage ||
      (fragment.physicalPage === best.physicalPage && fragment.y < best.y)
    ) {
      best = fragment;
    }
  }
  return best;
}

/** 建立版面索引。 */
export function buildLayoutIndex(layout: RawLayout): LayoutIndex {
  const pageByIndex = new Map<number, RawLayoutPage>();
  for (const page of layout.pages) pageByIndex.set(page.index, page);

  const fragmentsByBodyIndex = new Map<number, RawLayoutFragment[]>();
  const extentByPage = new Map<number, PageExtent>();

  for (const fragment of layout.fragments) {
    const list = fragmentsByBodyIndex.get(fragment.bodyIndex);
    if (list) list.push(fragment);
    else fragmentsByBodyIndex.set(fragment.bodyIndex, [fragment]);

    const key = fragment.physicalPage - 1;
    const current = extentByPage.get(key);
    const right = fragment.x + fragment.width;
    const bottom = fragment.y + fragment.height;
    if (current) {
      current.minX = Math.min(current.minX, fragment.x);
      current.minY = Math.min(current.minY, fragment.y);
      current.maxX = Math.max(current.maxX, right);
      current.maxY = Math.max(current.maxY, bottom);
    } else {
      extentByPage.set(key, { minX: fragment.x, minY: fragment.y, maxX: right, maxY: bottom });
    }
  }

  return { available: layout.available, unit: layout.unit, pageByIndex, fragmentsByBodyIndex, extentByPage };
}

/* -------------------------------------------------------------------------- */
/* 置信度与坐标                                                                */
/* -------------------------------------------------------------------------- */

function confidence(score: number, source: ConfidenceSource, notes: string[]): Confidence {
  return { score, source, notes };
}

/** 结构路径即稳定 id：可读、确定，且不依赖时间或随机数。 */
export function nodeIdFor(part: string, path: string): ComplexNodeId {
  return `${part}!${path}`;
}

function makeCoordinate(
  part: string,
  path: string,
  fragment: RawLayoutFragment | null,
  unit: CoordinateUnit,
): SourceCoordinate {
  if (!fragment) {
    return { part, structuralPath: path, pageIndex: null, box: null, unit: null };
  }
  return {
    part,
    structuralPath: path,
    pageIndex: fragment.physicalPage - 1,
    box: { x: fragment.x, y: fragment.y, width: fragment.width, height: fragment.height },
    unit,
  };
}

/* -------------------------------------------------------------------------- */
/* 映射                                                                        */
/* -------------------------------------------------------------------------- */

/** mapper 产出的告警。它只说「结果需要被注意」，不表示调用失败。 */
export interface MapperWarning {
  code: DocxComplexParseWarningCode;
  message: string;
  path?: string;
}

/** `execute` 归一化后的产物。 */
export interface MappedContent {
  content: ComplexContent;
  warnings: MapperWarning[];
}

/** 单元格的结构路径：由表格路径 + 原始行/格下标合成，与 OOXML 结构一致。 */
function cellPath(tablePath: string, row: number, cellIndex: number): string {
  return `${tablePath}/w:tr[${row + 1}]/w:tc[${cellIndex + 1}]`;
}

function toGridCells(
  part: string,
  tablePath: string,
  grid: LogicalGrid,
  fragment: RawLayoutFragment | null,
): GridCell[][] {
  // 版面坐标是表级的（引擎按块给几何），单元格沿用同一页号，但矩形留空——
  // 把整表的矩形冒充成单元格的，会比「不知道」更糟。

  return grid.grid.map((line) =>
    line.map((cell): GridCell => {
      const isReal = cell.origin !== null;
      const cellConfidence =
        cell.coverage === 'self'
          ? confidence(0.95, 'native', ['格子位置由 tblGrid 与 gridSpan 直接给出'])
          : cell.coverage === 'merged'
            ? confidence(0.9, 'structural', ['被左侧或上方的合并区域覆盖'])
            : confidence(0.4, 'heuristic', ['vMerge 续接但未找到起点']);

      return {
        row: cell.row,
        column: cell.column,
        text: cell.text,
        origin: cell.origin,
        coordinate: isReal
          ? {
              part,
              structuralPath: cellPath(tablePath, cell.row, cell.sourceIndex ?? 0),
              pageIndex: fragment ? fragment.physicalPage - 1 : null,
              box: null,
              unit: null,
            }
          : null,
        confidence: cellConfidence,
      };
    }),
  );
}

function toTableConfidence(grid: LogicalGrid, declared: number): Confidence {
  if (grid.issues.length > 0) {
    return confidence(0.6, 'structural', [
      '文档存在自相矛盾的合并声明，已按最保守方式解释',
      ...grid.issues.map((issue) => issue.message),
    ]);
  }
  if (declared <= 0) {
    return confidence(0.8, 'structural', ['tblGrid 缺失，列数由最长行推导']);
  }
  return confidence(0.98, 'native', ['格子由 tblGrid、gridSpan、vMerge 直接给出']);
}

function summarizeTable(grid: LogicalGrid): string {
  const lines: string[] = [];
  for (const line of grid.grid) {
    lines.push(line.map((cell) => cell.text).join('\t'));
  }
  return lines.join('\n');
}

function toComplexTable(
  id: ComplexNodeId,
  part: string,
  block: RawTableBlock,
  grid: LogicalGrid,
  coordinate: SourceCoordinate,
  fragment: RawLayoutFragment | null,
  fragments: readonly RawLayoutFragment[] | undefined,
): ComplexTable {
  const pages = new Set<number>();
  for (const item of fragments ?? []) pages.add(item.physicalPage - 1);
  const ordered = [...pages].sort((a, b) => a - b);

  return {
    id,
    coordinate,
    confidence: toTableConfidence(grid, block.tblGridColumns),
    columnCount: grid.columnCount,
    rowCount: grid.rowCount,
    // A cross-page table does not establish the page of each individual cell.
    grid: toGridCells(part, block.path, grid, ordered.length > 1 ? null : fragment),
    merges: grid.merges.map(
      (merge): MergeRegion => ({
        startRow: merge.startRow,
        startColumn: merge.startColumn,
        rowSpan: merge.rowSpan,
        columnSpan: merge.columnSpan,
        kind: merge.kind,
        text: merge.text,
      }),
    ),
    headerRowCount: Math.min(block.headerRowCount, grid.rowCount),
    // 版面可用时才算得出跨页；rdocx 按「块」给几何，跨页的表会给出多个片段。
    spansPages: fragment === null ? null : ordered.length > 1,
    pageRange: ordered.length === 0 ? null : [ordered[0] ?? 0, ordered[ordered.length - 1] ?? 0],
  };
}

function toPageInfos(
  layout: LayoutIndex,
  nodeIdsByPage: Map<number, ComplexNodeId[]>,
): PageInfo[] {
  const indices = new Set<number>(layout.pageByIndex.keys());
  for (const pageIndex of layout.extentByPage.keys()) indices.add(pageIndex);

  return [...indices]
    .sort((a, b) => a - b)
    .map((index): PageInfo => {
      const page = layout.pageByIndex.get(index) ?? null;
      const extent = layout.extentByPage.get(index) ?? null;

      if (!page) {
        // 有片段却无页尺寸：给出页号但明确 geometry 未知。
        return { index, geometry: null, nodeIds: nodeIdsByPage.get(index) ?? [], inferred: true };
      }

      // 页边距由片段外接范围反推——引擎只报页尺寸与块位置，不报 margin。
      const margins = extent
        ? {
            top: Math.max(0, extent.minY),
            right: Math.max(0, page.width - extent.maxX),
            bottom: Math.max(0, page.height - extent.maxY),
            left: Math.max(0, extent.minX),
          }
        : { top: 0, right: 0, bottom: 0, left: 0 };

      return {
        index,
        geometry: {
          width: page.width,
          height: page.height,
          unit: layout.unit,
          orientation: page.width >= page.height ? 'landscape' : 'portrait',
          margins,
        },
        nodeIds: nodeIdsByPage.get(index) ?? [],
        inferred: true,
      };
    });
}

function toBlock(
  part: string,
  raw: RawBlock,
  index: number,
  layout: LayoutIndex,
  warnings: MapperWarning[],
): { block: ComplexBlock; pageIndex: number | null } {
  const id = nodeIdFor(part, raw.path);
  const fragments = layout.available ? layout.fragmentsByBodyIndex.get(index) : undefined;
  const fragment = primaryFragment(fragments);
  const coordinate = makeCoordinate(part, raw.path, fragment, layout.unit);

  if (raw.kind === 'paragraph') {
    return {
      block: {
        id,
        kind: 'paragraph',
        text: raw.text,
        coordinate,
        confidence: confidence(0.95, 'native', ['段落文本由 OOXML 直接给出']),
        table: null,
      },
      pageIndex: coordinate.pageIndex,
    };
  }

  const grid = expandLogicalGrid(raw);
  for (const issue of grid.issues) {
    warnings.push({ code: 'MERGE_INCONSISTENT', message: issue.message, path: raw.path });
  }

  return {
    block: {
      id,
      kind: 'table',
      text: summarizeTable(grid),
      coordinate,
      confidence: toTableConfidence(grid, raw.tblGridColumns),
      table: toComplexTable(id, part, raw, grid, coordinate, fragment, fragments),
    },
    pageIndex: coordinate.pageIndex,
  };
}

/**
 * 把引擎的原始观察归一化为本模块的复杂结构载荷。
 *
 * 纯函数：同样的 `raw` 必然产出逐字段相同的结果，便于快照测试与缓存。
 */
export function toComplexContent(raw: RawParseResult, flags: Partial<FeatureFlags> = {}): MappedContent {
  // 引擎上报的告警先于 mapper 自己的告警入队：前者说的是「读取时看到了什么」，
  // 后者说的是「归一化时判断出了什么」，按因果顺序排列，宿主读起来更顺。
  const warnings: MapperWarning[] = raw.warnings.map((warning) =>
    warning.path === undefined
      ? { code: warning.code, message: warning.message }
      : { code: warning.code, message: warning.message, path: warning.path },
  );
  const layout = buildLayoutIndex(raw.layout.available && flags.parsePageGeometry !== false
    ? raw.layout : { available: false, unit: raw.layout.unit, pages: [], fragments: [] });

  const blocks: ComplexBlock[] = [];
  const nodeIdsByPage = new Map<number, ComplexNodeId[]>();

  for (let i = 0; i < raw.blocks.length; i += 1) {
    const rawBlock = raw.blocks[i];
    if (!rawBlock) continue;

    const { block } = toBlock(raw.mainPart, rawBlock, i, layout, warnings);
    const fragments = layout.fragmentsByBodyIndex.get(i) ?? [];
    block.fragments = fragments.map((fragment) =>
      makeCoordinate(raw.mainPart, rawBlock.path, fragment, layout.unit));
    blocks.push(block);

    for (const pageIndex of new Set(fragments.map((fragment) => fragment.physicalPage - 1))) {
      const list = nodeIdsByPage.get(pageIndex);
      if (list) list.push(block.id);
      else nodeIdsByPage.set(pageIndex, [block.id]);
    }
  }

  const breaks: PageBreak[] = flags.parsePageBreaks === false ? [] : (raw.breaks ?? []).map((item) => {
    const before = item.beforeBodyIndex === null ? null : blocks[item.beforeBodyIndex];
    return {
      kind: item.kind, beforeNodeId: before?.id ?? null,
      // A rendered marker records a past layout, not a trustworthy current page number.
      pageIndexBefore: null, pageIndexAfter: null,
      coordinate: makeCoordinate(raw.mainPart, item.path, null, layout.unit),
      confidence: confidence(1, 'native', [item.kind === 'rendered'
        ? 'Saved rendering marker; current page numbers are unknown.'
        : 'Break declaration read directly from OOXML.']),
    };
  });
  const sections = flags.parseSections === false ? [] : (raw.sections ?? []).map((item, index) => ({
    index,
    startBeforeNodeId: item.startBodyIndex === null ? null : blocks[item.startBodyIndex]?.id ?? null,
    geometry: item.geometry === null ? null : { ...item.geometry, margins: { ...item.geometry.margins } },
    columnCount: item.columnCount,
    startType: item.startType,
  }));
  const floats: FloatingObject[] = flags.parseFloatingObjects === false ? [] : (raw.floats ?? []).map((item) => {
    warnings.push({ code: 'FLOAT_ANCHOR_APPROXIMATED', path: item.path,
      message: 'Drawing anchor is preserved; object-specific page and rectangle are unavailable.' });
    return {
      id: nodeIdFor(raw.mainPart, item.path), kind: item.kind,
      coordinate: makeCoordinate(raw.mainPart, item.path, null, layout.unit),
      confidence: confidence(1, 'native', ['DrawingML declaration; absolute layout is not inferred from its parent block.']),
      anchor: { ...item.anchor }, width: item.width, height: item.height, unit: item.unit,
      part: raw.mainPart, relationshipId: item.relationshipId, name: item.name, altText: item.altText,
    };
  });
  const coverage = (enabled: boolean | undefined, collection: unknown[] | undefined):
    'observed' | 'disabled' | 'unavailable' => enabled === false ? 'disabled'
      : collection === undefined ? 'unavailable' : 'observed';

  if (blocks.length === 0) {
    warnings.push({ code: 'EMPTY_DOCUMENT', message: 'No visible body blocks were found.' });
  }
  if (!layout.available && raw.blocks.length > 0) {
    warnings.push({
      code: 'PAGE_GEOMETRY_UNAVAILABLE',
      message: 'Page geometry is unavailable; coordinates degrade to structural paths.',
    });
  }

  return {
    content: {
      scheme: COMPLEX_LAYOUT_SCHEME,
      mainPart: raw.mainPart,
      pages: toPageInfos(layout, nodeIdsByPage),
      blocks,
      breaks,
      sections,
      floats,
      coverage: {
        breaks: coverage(flags.parsePageBreaks, raw.breaks),
        sections: coverage(flags.parseSections, raw.sections),
        floats: flags.parseFloatingObjects !== false && raw.floats !== undefined
          && raw.warnings.some((warning) => warning.code === 'UNSUPPORTED_CONTENT')
          ? 'partial' : coverage(flags.parseFloatingObjects, raw.floats),
      },
    },
    warnings,
  };
}

/* -------------------------------------------------------------------------- */
/* 包级事实 → FormatProfile / FormatIR                                          */
/* -------------------------------------------------------------------------- */

/**
 * 格式判定的置信度：每个独立信号各占一份权重。
 *
 * 之所以按信号加权而不是「识别到了就打 1.0」：magic 说这是 zip、Content-Types
 * 说这是 WordprocessingML、部件名说主文档在这儿——三条互相独立，越多条印证
 * 越不容易被改名/伪装的文件骗到。缺一条就该掉一档，而不是四舍五入到满分。
 */
function formatConfidence(signatures: readonly string[]): number {
  let score = 0;
  if (signatures.includes('magic:zip')) score += 0.4;
  if (signatures.includes('content-types:wordprocessingml')) score += 0.4;
  if (signatures.some((signature) => signature.startsWith('part:'))) score += 0.2;
  return score;
}

/** 生成 `FormatProfile` 所需的外部输入。 */
export interface FormatProfileInput {
  /** 依据 artifact 引用得到的扩展名（小写、不含点）；无法判定时为 null。 */
  extension: string | null;
}

/**
 * 包级原始事实 → `FormatProfile`。
 *
 * `features` 只声明包级可确认的指标。DrawingML 扫描仍不覆盖所有部件与 VML，
 * 因此不能将正文扫描的空结果写成全包 `hasFloatingObjects: false`；内容侧
 * 通过 coverage 明确表达其观察范围与可用性。
 */
export function toFormatProfile(raw: RawParseResult, input: FormatProfileInput): FormatProfile {
  const tableCount = raw.blocks.filter((block) => block.kind === 'table').length;

  return {
    format: raw.profile.format,
    mediaType: raw.profile.mediaType,
    extension: input.extension,
    confidence: formatConfidence(raw.profile.signatures),
    container: raw.profile.container,
    encrypted: raw.profile.encrypted,
    signatures: [...raw.profile.signatures],
    features: {
      hasMacros: raw.indicators.macros.length > 0,
      hasExternalReferences: raw.indicators.externalReferences.length > 0,
      hasEmbeddedObjects: raw.indicators.embeddedObjects.length > 0,
      hasPageGeometry: raw.layout.available,
    },
    metadata: {
      parseVersion: raw.parseVersion,
      mainPart: raw.mainPart,
      partCount: raw.parts.length,
      relationshipCount: raw.relationships.length,
      blockCount: raw.blocks.length,
      tableCount,
      pageCount: raw.layout.pages.length,
      pageGeometryUnit: raw.layout.unit,
    },
  };
}

/**
 * 包级事实 + 归一化内容 → 完整 IR。
 *
 * `parts` / `relationships` / `indicators` 与 office-core 的同名类型逐字段同构，
 * 因此直接透传：多写一层逐字段拷贝，只会在字段增减时引入一处忘改的地方。
 */
export function toComplexIR(
  raw: RawParseResult,
  content: ComplexContent,
  profile: FormatProfile,
): DocxComplexParseIR {
  return {
    profile,
    parts: raw.parts,
    relationships: raw.relationships,
    indicators: raw.indicators,
    content,
  };
}

/* -------------------------------------------------------------------------- */
/* 告警严重级别                                                                */
/* -------------------------------------------------------------------------- */

/**
 * 告警码 → 严重级别。
 *
 * 写死成表而不是「按消息猜」：编排层可能只消费 `severity`，让它随文案摇摆，
 * 等于把安全判断交给措辞。缺键会编译报错，因此新增告警码不会漏配。
 */
const WARNING_SEVERITY: Record<DocxComplexParseWarningCode, WarningSeverity> = {
  EMPTY_DOCUMENT: 'info',
  PARTIALLY_PARSED: 'warn',
  UNSUPPORTED_CONTENT: 'warn',
  LIMIT_APPLIED: 'warn',
  NESTED_TABLE_FLATTENED: 'warn',
  MERGE_INCONSISTENT: 'warn',
  PAGE_GEOMETRY_UNAVAILABLE: 'warn',
  PAGE_BREAK_ESTIMATED: 'info',
  FLOAT_ANCHOR_APPROXIMATED: 'warn',
  LOW_CONFIDENCE: 'info',
  ENCRYPTED_ARTIFACT: 'error',
  VERIFICATION_FAILED: 'error',
};

/** mapper 告警 → 面向编排层的告警。 */
export function toWarnings(warnings: readonly MapperWarning[]): Warning[] {
  return warnings.map((warning) => {
    const base: Warning = {
      code: warning.code,
      message: warning.message,
      severity: WARNING_SEVERITY[warning.code],
    };
    return warning.path === undefined ? base : { ...base, path: warning.path };
  });
}

/* -------------------------------------------------------------------------- */
/* 预算与低置信度                                                              */
/* -------------------------------------------------------------------------- */

/** 统计全部表格的单元格数量（逻辑网格的格子数，含被合并覆盖的占位格）。 */
export function countTableCells(content: ComplexContent): number {
  let total = 0;
  for (const block of content.blocks) {
    if (block.table) total += block.table.rowCount * block.table.columnCount;
  }
  return total;
}

/** 一个结构上的置信度与其定位信息。 */
interface ConfidenceSite {
  score: number;
  path: string;
}

function collectConfidenceSites(content: ComplexContent): ConfidenceSite[] {
  const sites: ConfidenceSite[] = [];
  for (const block of content.blocks) {
    sites.push({ score: block.confidence.score, path: block.coordinate.structuralPath });
    if (!block.table) continue;
    for (const line of block.table.grid) {
      for (const cell of line) {
        if (!cell.coordinate) continue;
        sites.push({ score: cell.confidence.score, path: cell.coordinate.structuralPath });
      }
    }
  }
  for (const item of [...content.breaks, ...content.floats]) {
    sites.push({ score: item.confidence.score, path: item.coordinate.structuralPath });
  }
  return sites;
}

/**
 * 单次最多广播多少条 `LOW_CONFIDENCE`。
 *
 * 一份整表都是续接格却找不到起点的文档，能产生上万条低置信度结论。逐条广播会把
 * 告警列表淹没，反而让上层漏掉真正重要的那几条——所以超出部分只汇总报一次。
 */
const MAX_LOW_CONFIDENCE_WARNINGS = 50;

/** 找出低于置信度下限的结论，广播 `LOW_CONFIDENCE`。 */
export function collectLowConfidence(
  content: ComplexContent,
  floor: number,
): MapperWarning[] {
  const low = collectConfidenceSites(content).filter((site) => site.score < floor);
  if (low.length === 0) return [];

  const warnings: MapperWarning[] = low.slice(0, MAX_LOW_CONFIDENCE_WARNINGS).map((site) => ({
    code: 'LOW_CONFIDENCE' as const,
    message: `Confidence ${site.score} is below the floor ${floor}.`,
    path: site.path,
  }));

  if (low.length > MAX_LOW_CONFIDENCE_WARNINGS) {
    warnings.push({
      code: 'LOW_CONFIDENCE',
      message:
        `${low.length - MAX_LOW_CONFIDENCE_WARNINGS} more conclusion(s) fell below ` +
        `the confidence floor ${floor}; only the first ${MAX_LOW_CONFIDENCE_WARNINGS} ` +
        'are listed individually.',
    });
  }
  return warnings;
}

/* -------------------------------------------------------------------------- */
/* 可用性与资源预算                                                            */
/* -------------------------------------------------------------------------- */

/** 一项被突破的资源预算。 */
export interface BudgetOverrun {
  limit: keyof LimitConfig;
  /** 实际用量。 */
  value: number;
  /** 配置的上限。 */
  max: number;
}

/**
 * 量出各项预算的实际用量。
 *
 * Optional observation collections are measured only when provided by the engine.
 */
function measureBudgets(raw: RawParseResult): { limit: keyof LimitConfig; value: number }[] {
  let tableCells = 0;
  let uncompressedBytes = 0;
  let largestPart = 0;
  for (const part of raw.parts) {
    uncompressedBytes += part.sizeBytes;
    if (part.sizeBytes > largestPart) largestPart = part.sizeBytes;
  }
  for (const block of raw.blocks) {
    if (block.kind !== 'table') continue;
    const grid = expandLogicalGrid(block);
    tableCells += grid.rowCount * grid.columnCount;
  }

  return [
    { limit: 'maxArchiveEntries', value: raw.parts.length },
    { limit: 'maxEntryUncompressedBytes', value: largestPart },
    { limit: 'maxTotalUncompressedBytes', value: uncompressedBytes },
    { limit: 'maxRelationships', value: raw.relationships.length },
    { limit: 'maxBlocks', value: raw.blocks.length },
    { limit: 'maxTableCells', value: tableCells },
    { limit: 'maxPages', value: raw.layout.pages.length },
    ...(raw.floats === undefined ? [] : [{ limit: 'maxFloatingObjects' as const, value: raw.floats.length }]),
  ];
}

/**
 * 判断这份观察结果能不能当作 DOCX 继续处理，并量出预算超支。
 *
 * 返回超支清单而不是直接失败：`enforceLimits` 关闭时调用方需要把它降级成
 * `LIMIT_APPLIED` 告警。也就是说这里的返回值是「事实」，抛不抛是配置说了算。
 *
 * @throws DocxComplexParseError 格式不符、容器不可读、加密，或限额开启且已超支。
 */
export function assertParseUsable(raw: RawParseResult, config: ModuleConfig): BudgetOverrun[] {
  const { profile } = raw;

  if (profile.container === 'ole') {
    throw new DocxComplexParseError(
      'UNSUPPORTED_CONTAINER',
      'The artifact is an OLE compound file (encrypted, or a legacy .doc).',
      { details: { container: profile.container } },
    );
  }
  if (profile.container !== 'zip') {
    throw new DocxComplexParseError(
      'FORMAT_MISMATCH',
      `The artifact container is "${profile.container}", not a DOCX package.`,
      { details: { container: profile.container, format: profile.format } },
    );
  }
  if (profile.encrypted === true) {
    throw new DocxComplexParseError(
      'UNSUPPORTED_CONTAINER',
      'The artifact is encrypted and cannot be read.',
      { details: { encrypted: true } },
    );
  }
  if (profile.format === 'unknown') {
    throw new DocxComplexParseError(
      'FORMAT_MISMATCH',
      'The package is not a WordprocessingML document.',
      { details: { mediaType: profile.mediaType, signatures: profile.signatures } },
    );
  }
  // 调用方说「没有版面就当失败」时，宁可在这里明确拒绝，也不要交出一份
  // 坐标全为 null 的结果——后者看起来像成功。
  if (config.featureFlags.requirePageGeometry && !raw.layout.available) {
    throw new DocxComplexParseError(
      'LAYOUT_UNAVAILABLE',
      'Page geometry is required by the active feature flags but unavailable.',
    );
  }

  const overruns: BudgetOverrun[] = [];
  for (const { limit, value } of measureBudgets(raw)) {
    const max = config.limits[limit];
    if (max > 0 && value > max) overruns.push({ limit, value, max });
  }

  const first = overruns[0];
  if (config.featureFlags.enforceLimits && first) {
    throw new DocxComplexParseError(
      'LIMIT_EXCEEDED',
      `${first.limit} was exceeded: ${first.value} > ${first.max}.`,
      { details: { limit: first.limit, value: first.value, max: first.max } },
    );
  }
  return overruns;
}

/**
 * 补全 artifact 引用。
 *
 * 只回填我们【确实知道】的字段（媒体类型）。字节数与摘要属于容器级事实，
 * 本模块不读原始字节流，因此宁可留空，也不能算一个「差不多」的数字出来。
 */
export function refineArtifactRef(ref: ArtifactRef, raw: RawParseResult): ArtifactRef {
  if (ref.mediaType !== undefined) return { ...ref };
  const mediaType = raw.profile.mediaType;
  return mediaType === null ? { ...ref } : { ...ref, mediaType };
}

/* -------------------------------------------------------------------------- */
/* 调用方的类型声明 vs 实际内容                                                 */
/* -------------------------------------------------------------------------- */

/** 扩展名 → 格式身份。不在这张表里的扩展名不做判断。 */
const FORMAT_BY_EXTENSION: Record<string, RawFormatKind> = {
  docx: 'docx',
  docm: 'docm',
  dotx: 'dotx',
  dotm: 'dotm',
};

/**
 * 核对调用方声称的扩展名/媒体类型与实际字节内容。
 *
 * 只在**两侧都已知且冲突**时失败：调用方没声明、或声明了一个我们不认的扩展名，
 * 都不构成「内容不对」的证据。这条纪律避免把「.DOCX 这种大写扩展名」之类的
 * 无害差异误报成类型错误。
 *
 * 之所以宁可信字节也不信扩展名：宏就是靠改扩展名绕过检查的（把 .docm 改名成
 * .docx）。这里失败能挡住这类伪装——但注意，真正的防线是 `indicators.macros`，
 * 本检查只是补充。
 *
 * @throws DocxComplexParseError 声明的类型与内容不符时抛出 `FORMAT_MISMATCH`。
 */
export function assertDeclaredTypeMatches(
  raw: RawParseResult,
  declared: { extension?: string | undefined; mimeType?: string | undefined },
): void {
  const extension = declared.extension?.trim().replace(/^\./, '').toLowerCase();
  if (extension !== undefined && extension.length > 0) {
    const claimed = FORMAT_BY_EXTENSION[extension];
    if (claimed !== undefined && claimed !== raw.profile.format) {
      throw new DocxComplexParseError(
        'FORMAT_MISMATCH',
        `The caller declared a "${extension}" artifact but the content is "${raw.profile.format}".`,
        { details: { declaredExtension: extension, actualFormat: raw.profile.format } },
      );
    }
  }

  const mimeType = declared.mimeType?.trim().toLowerCase();
  // mediaType 为 null 表示引擎没能从 Content-Types 里认出主文档类型，
  // 此时「声明」与「未知」不构成冲突。
  if (mimeType !== undefined && mimeType.length > 0 && raw.profile.mediaType !== null) {
    if (mimeType !== raw.profile.mediaType.toLowerCase()) {
      throw new DocxComplexParseError(
        'FORMAT_MISMATCH',
        `The declared media type "${declared.mimeType}" does not match the content.`,
        { details: { declaredMimeType: declared.mimeType, actualMediaType: raw.profile.mediaType } },
      );
    }
  }
}

/** 从 artifact 引用推扩展名；拿不到时返回 null。 */
export function extensionOf(ref: ArtifactRef): string | null {
  const label = ref.label ?? ref.uri;
  // 去掉查询串与片段，再取最后一个点之后的部分。
  const withoutQuery = label.split(/[?#]/)[0] ?? label;
  const base = withoutQuery.split(/[\\/]/).pop() ?? '';
  const dot = base.lastIndexOf('.');
  if (dot <= 0 || dot === base.length - 1) return null;
  return base.slice(dot + 1).toLowerCase();
}

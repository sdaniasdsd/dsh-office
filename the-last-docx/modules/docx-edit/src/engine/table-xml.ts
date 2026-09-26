/**
 * Table instance geometry.
 *
 * A table style only says what a table *may* look like. What decides the
 * layout is the table element itself: `w:tblW`, `w:tblLayout`, the `w:gridCol`
 * widths in `w:tblGrid`, and each cell's `w:tcW`. A `<w:tbl>` with an empty
 * `<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>` - which is what a
 * minimally-built document carries - has **no preferred widths at all**, so a
 * renderer falls back to the minimum width the content needs and `A1` wraps to
 * two lines. Defining a table style cannot fix that; the geometry has to be
 * written on the table.
 *
 * Every mature OOXML library keeps this on the content object rather than in
 * the style collection for the same reason: python-docx exposes
 * `table.autofit` / `column.width` / `cell.width`, Apache POI exposes
 * `XWPFTable.setWidth`, and docx4j exposes the `Tbl` itself. docx4j-core-ts
 * gives no writable accessor for `w:tblW`, `w:tblGrid` or `w:tblCellMar`, so
 * the tree is edited directly here.
 *
 * Units are points, as everywhere else in this module; the engine converts to
 * twips (`pt * 20`).
 */
import type { Element, Node } from '@xmldom/xmldom';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxEditError } from '../errors';
import type { TableGeometry } from '../domain/docx-edit';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const toTwips = (points: number): number => Math.round(points * 20);

export type { TableBorderEdge, TableBorderEdgeSpec, TableGeometry } from '../domain/docx-edit';

/** Child sequence ECMA-376 declares for `w:tblPr`. */
const TBLPR_ORDER = [
  'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
  'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
] as const;
const TBLBORDERS_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'] as const;

export function parseDocumentPart(source: string) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) {
    throw new DocxEditError('SAFETY_POLICY_DENIED', 'DTD and entity declarations are forbidden.');
  }
  try {
    return new DOMParser({ onError: () => { throw new Error('Malformed XML'); } }).parseFromString(source, 'application/xml');
  } catch {
    throw new DocxEditError('FORMAT_MISMATCH', 'The main document part is malformed.');
  }
}

export const serializeDocumentPart = (document: ReturnType<typeof parseDocumentPart>): string =>
  new XMLSerializer().serializeToString(document);

const localName = (element: Element): string => element.localName ?? element.nodeName.replace(/^.*:/, '');

function children(element: Element): Element[] {
  const out: Element[] = [];
  const nodes = element.childNodes;
  for (let index = 0; index < nodes.length; index += 1) {
    const node: Node | undefined = nodes.item(index) ?? undefined;
    if (node && node.nodeType === 1) out.push(node as Element);
  }
  return out;
}

const childNamed = (element: Element, name: string): Element | undefined =>
  children(element).find((child) => localName(child) === name);

/** Every `w:tbl` reachable in document order, depth-first, nesting included. */
function tablesIn(root: Element): Element[] {
  const out: Element[] = [];
  const walk = (element: Element): void => {
    for (const child of children(element)) {
      if (localName(child) === 'tbl') out.push(child);
      walk(child);
    }
  };
  walk(root);
  return out;
}

/** Replace `container`'s managed children with `wanted`, keeping the declared order. */
function setOrderedChildren(document: ReturnType<typeof parseDocumentPart>, container: Element, order: readonly string[], wanted: Record<string, string | undefined>): void {
  for (const name of order) {
    const existing = childNamed(container, name);
    const markup = wanted[name];
    if (existing && markup === undefined) continue;
    if (existing) container.removeChild(existing);
    if (markup !== undefined) container.appendChild(parseDocumentPart(`<w:root xmlns:w="${W}">${markup}</w:root>`).documentElement!.firstChild as Node);
  }
  // Re-emit in the declared sequence so the order is right by construction.
  const present = children(container).filter((child) => order.includes(localName(child)));
  const rank = new Map(order.map((name, index) => [name, index]));
  present.sort((a, b) => (rank.get(localName(a)) ?? 0) - (rank.get(localName(b)) ?? 0));
  for (const child of present) container.appendChild(child);
  void document;
}

function ensureChild(document: ReturnType<typeof parseDocumentPart>, parent: Element, name: string, first: boolean): Element {
  const existing = childNamed(parent, name);
  if (existing) return existing;
  const created = document.createElementNS(W, `w:${name}`);
  if (first && parent.firstChild) parent.insertBefore(created, parent.firstChild);
  else parent.appendChild(created);
  return created;
}

/**
 * Write the geometry onto the `index`-th table of the main document part.
 *
 * Returns the rewritten part and the property names that were actually
 * applied, so the caller can verify what it asked for instead of assuming it.
 */
export function applyTableGeometry(mainXml: string, index: number, geometry: TableGeometry): { xml: string; applied: string[]; columnCount: number } {
  const document = parseDocumentPart(mainXml);
  const root = document.documentElement as Element | null;
  if (!root) throw new DocxEditError('FORMAT_MISMATCH', 'The main document part has no root element.');
  const table = tablesIn(root)[index];
  if (!table) throw new DocxEditError('TARGET_NOT_FOUND', `The document has no table at index ${index}.`);

  const applied: string[] = [];
  const tblPr = ensureChild(document, table, 'tblPr', true);
  const wanted: Record<string, string | undefined> = {};

  if (geometry.styleId !== undefined) {
    wanted.tblStyle = `<w:tblStyle w:val="${escapeAttr(geometry.styleId)}"/>`;
    applied.push('styleId');
  }
  if (geometry.width !== undefined) {
    wanted.tblW = `<w:tblW w:w="${toTwips(geometry.width)}" w:type="dxa"/>`;
    applied.push('width');
  }
  if (geometry.alignment !== undefined) {
    wanted.jc = `<w:jc w:val="${escapeAttr(geometry.alignment)}"/>`;
    applied.push('alignment');
  }
  if (geometry.borders !== undefined) {
    const edges = TBLBORDERS_ORDER
      .filter((edge) => geometry.borders?.[edge] !== undefined)
      .map((edge) => {
        const spec = geometry.borders![edge]!;
        return `<w:${edge} w:val="${escapeAttr(spec.style ?? 'single')}" w:sz="${Math.round((spec.size ?? 0.5) * 8)}" w:space="0" w:color="${escapeAttr(colorOf(spec.color))}"/>`;
      });
    wanted.tblBorders = edges.length > 0 ? `<w:tblBorders>${edges.join('')}</w:tblBorders>` : undefined;
    if (edges.length > 0) applied.push('borders');
  }
  if (geometry.layout !== undefined) {
    wanted.tblLayout = `<w:tblLayout w:type="${geometry.layout === 'fixed' ? 'fixed' : 'autofit'}"/>`;
    applied.push('layout');
  }
  if (geometry.cellMargins !== undefined) {
    const margins = (['top', 'left', 'bottom', 'right'] as const)
      .filter((edge) => geometry.cellMargins?.[edge] !== undefined)
      .map((edge) => `<w:${edge} w:w="${toTwips(geometry.cellMargins![edge]!)}" w:type="dxa"/>`);
    wanted.tblCellMar = margins.length > 0 ? `<w:tblCellMar>${margins.join('')}</w:tblCellMar>` : undefined;
    if (margins.length > 0) applied.push('cellMargins');
  }
  setOrderedChildren(document, tblPr, TBLPR_ORDER, wanted);

  // The grid carries the preferred column widths; without them a renderer sizes
  // the table to its content and long cell text wraps.
  const rows = children(table).filter((child) => localName(child) === 'tr');
  const firstRowCells = rows[0] ? children(rows[0]).filter((child) => localName(child) === 'tc') : [];
  const columnCount = firstRowCells.length;
  if (geometry.columnWidths !== undefined) {
    if (geometry.columnWidths.length !== columnCount) {
      throw new DocxEditError('INVALID_INPUT', `columnWidths has ${geometry.columnWidths.length} entries but the table has ${columnCount} columns.`);
    }
    const grid = ensureChild(document, table, 'tblGrid', false);
    for (const existing of children(grid)) grid.removeChild(existing);
    for (const width of geometry.columnWidths) {
      const column = document.createElementNS(W, 'w:gridCol');
      column.setAttribute('w:w', String(toTwips(width)));
      grid.appendChild(column);
    }
    // The grid must follow tblPr, so re-seat it if anything landed after it.
    if (grid.previousSibling !== tblPr) table.insertBefore(grid, tblPr.nextSibling);
    applied.push('columnWidths');
  }

  if (geometry.headerRow !== undefined || geometry.cellVerticalAlignment !== undefined) {
    for (const [rowIndex, row] of rows.entries()) {
      if (geometry.headerRow !== undefined && rowIndex === 0) {
        const trPr = ensureChild(document, row, 'trPr', true);
        const existing = childNamed(trPr, 'tblHeader');
        if (existing) trPr.removeChild(existing);
        if (geometry.headerRow) {
          const header = document.createElementNS(W, 'w:tblHeader');
          trPr.appendChild(header);
        }
        if (rowIndex === 0) applied.push('headerRow');
      }
      if (geometry.cellVerticalAlignment === undefined) continue;
      for (const cell of children(row)) {
        if (localName(cell) !== 'tc') continue;
        const tcPr = ensureChild(document, cell, 'tcPr', true);
        setOrderedChildren(document, tcPr, TCPR_ORDER, {
          vAlign: `<w:vAlign w:val="${escapeAttr(geometry.cellVerticalAlignment)}"/>`,
        });
      }
    }
    if (geometry.cellVerticalAlignment !== undefined) applied.push('cellVerticalAlignment');
  }

  if (geometry.columnWidths !== undefined) {
    // Cell widths have to agree with the grid, or Word re-derives the layout.
    for (const row of rows) {
      for (const [cellIndex, cell] of children(row).filter((child) => localName(child) === 'tc').entries()) {
        const width = geometry.columnWidths[cellIndex];
        if (width === undefined) continue;
        const tcPr = ensureChild(document, cell, 'tcPr', true);
        setOrderedChildren(document, tcPr, TCPR_ORDER, {
          tcW: `<w:tcW w:w="${toTwips(width)}" w:type="dxa"/>`,
        });
      }
    }
  }

  return { xml: serializeDocumentPart(document), applied, columnCount };
}

const escapeAttr = (value: string): string => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const colorOf = (color: string | undefined): string =>
  color === undefined ? 'auto' : /^#[0-9a-fA-F]{6}$/.test(color) ? color.slice(1).toUpperCase() : color;

/** What a table in a saved part actually carries. */
export interface TableGeometryReading {
  present: boolean;
  columnCount: number;
  /** Rows of cells, so a per-column expectation can be compared row by row. */
  rowCount: number;
  styleId?: string;
  widthTwips?: number;
  layout?: string;
  /** One entry per column; 0 means the grid column declared no preferred width. */
  gridWidthsTwips: number[];
  hasCellMargins: boolean;
  hasBorders: boolean;
  headerRow: boolean;
  cellWidthsTwips: number[];
}

/**
 * Read a table's geometry back out of a saved part.
 *
 * The property that matters most is `gridWidthsTwips`: a bare `<w:gridCol/>`
 * reads 0, and a table whose every column reads 0 is exactly the table that
 * wraps its cell text.
 */
export function readTableGeometry(mainXml: string, index: number): TableGeometryReading {
  const document = parseDocumentPart(mainXml);
  const root = document.documentElement as Element | null;
  const table = root ? tablesIn(root)[index] : undefined;
  if (!table) {
    return { present: false, columnCount: 0, rowCount: 0, gridWidthsTwips: [], hasCellMargins: false, hasBorders: false, headerRow: false, cellWidthsTwips: [] };
  }
  const tblPr = childNamed(table, 'tblPr');
  const readTwips = (element: Element | undefined, attribute: string): number => {
    const raw = element?.getAttribute(attribute);
    const value = raw ? Number.parseInt(raw, 10) : Number.NaN;
    return Number.isFinite(value) ? value : 0;
  };
  const grid = childNamed(table, 'tblGrid');
  const rows = children(table).filter((child) => localName(child) === 'tr');
  const firstRow = rows[0];
  const trPr = firstRow ? childNamed(firstRow, 'trPr') : undefined;
  const reading: TableGeometryReading = {
    present: true,
    columnCount: (grid ? children(grid).filter((child) => localName(child) === 'gridCol') : []).length,
    rowCount: rows.length,
    gridWidthsTwips: (grid ? children(grid).filter((child) => localName(child) === 'gridCol') : []).map((column) => readTwips(column, 'w:w')),
    hasCellMargins: tblPr ? childNamed(tblPr, 'tblCellMar') !== undefined : false,
    hasBorders: tblPr ? childNamed(tblPr, 'tblBorders') !== undefined : false,
    headerRow: trPr ? childNamed(trPr, 'tblHeader') !== undefined : false,
    cellWidthsTwips: rows.flatMap((row) => children(row)
      .filter((child) => localName(child) === 'tc')
      .map((cell) => readTwips(childNamed(cell, 'tcPr') ? childNamed(childNamed(cell, 'tcPr')!, 'tcW') : undefined, 'w:w'))),
  };
  const style = tblPr ? childNamed(tblPr, 'tblStyle') : undefined;
  const styleId = style?.getAttribute('w:val');
  if (styleId) reading.styleId = styleId;
  const width = tblPr ? childNamed(tblPr, 'tblW') : undefined;
  if (width) reading.widthTwips = readTwips(width, 'w:w');
  const layout = tblPr ? childNamed(tblPr, 'tblLayout') : undefined;
  const layoutType = layout?.getAttribute('w:type');
  if (layoutType) reading.layout = layoutType;
  return reading;
}

/** Validate a geometry request before the document is touched. */
export function assertTableGeometry(geometry: TableGeometry): void {
  const named = Object.keys(geometry)
    .filter((key) => key !== 'kind' && key !== 'target')
    .filter((key) => (geometry as Record<string, unknown>)[key] !== undefined);
  if (named.length === 0) {
    throw new DocxEditError('INVALID_INPUT', 'A formatTable edit must set at least one property.');
  }
  if (geometry.width !== undefined && (!Number.isFinite(geometry.width) || geometry.width <= 0)) {
    throw new DocxEditError('INVALID_INPUT', 'width must be a positive number of points.');
  }
  for (const width of geometry.columnWidths ?? []) {
    if (!Number.isFinite(width) || width <= 0) {
      throw new DocxEditError('INVALID_INPUT', 'Every columnWidths entry must be a positive number of points.');
    }
  }
  if (geometry.width !== undefined && geometry.columnWidths !== undefined) {
    const total = geometry.columnWidths.reduce((sum, width) => sum + width, 0);
    if (Math.abs(total - geometry.width) > 1) {
      throw new DocxEditError('INVALID_INPUT', `columnWidths totals ${total.toFixed(1)}pt but width is ${geometry.width}pt; a fixed-layout grid must agree with the table width.`);
    }
  }
  if (geometry.styleId !== undefined && geometry.styleId.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'styleId must be a non-empty string.');
  }
  for (const [edge, spec] of Object.entries(geometry.borders ?? {})) {
    if (spec?.color !== undefined && !/^(#[0-9a-fA-F]{6}|auto)$/.test(spec.color)) {
      throw new DocxEditError('INVALID_INPUT', `borders.${edge}.color must be '#RRGGBB' or 'auto'.`);
    }
    if (spec?.size !== undefined && (!Number.isFinite(spec.size) || spec.size < 0)) {
      throw new DocxEditError('INVALID_INPUT', `borders.${edge}.size must be a non-negative number of points.`);
    }
  }
}

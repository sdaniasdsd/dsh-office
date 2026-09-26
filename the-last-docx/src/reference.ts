/**
 * Reference-driven creation: turn a parsed document into a creation plan.
 *
 * Why this exists
 * ---------------
 * `docx-create` could only build from a plan an author wrote by hand, and
 * `docx-parse` could describe a document in full, but nothing joined the two: a
 * caller holding a reference document had no way to say "produce this again".
 * The only bridge the profile had, `docx_analyze`, produced evidence for a
 * reader, not a plan for the writing engine.
 *
 * What this is not
 * ----------------
 * Not a lossless reformat. A document built from a plan is a new document:
 * comments, footnotes, hyperlinks, bookmarks, revision marks, images and field
 * codes have no expression in a creation plan and are **dropped**, which the
 * returned report states rather than hides. Reformatting a document that carries
 * annotations is `docx-edit`'s job, on the document itself.
 */
import type { CreatePlan, DesignRegister, Preset, Scenario } from '@dsh-office-profile/docx-create';

/**
 * The part of docx-parse's dual IR this mapper reads.
 *
 * Declared structurally rather than imported: the mapper depends on the fields it
 * actually consumes, so an unrelated change to the parse module's types cannot
 * break it, and the fields it needs are written down here instead of being
 * inferred from a type that carries far more. A drift in these fields is caught
 * by `tests/reference.test.ts`, which runs against a real parse.
 */
export interface SemanticContent {
  semantic: {
    blocks: SemanticBlock[];
    annotations?: { kind: string }[];
  };
}

export interface SemanticBlock {
  kind: string;
  text?: string;
  styleId?: string | null;
  /** Headings only: 1-based outline level. */
  level?: number;
  /** List items only: which numbering definition the item belongs to, and its level. */
  list?: { numId?: string; level?: number } | null;
  /** Tables only. */
  gridColumns?: number;
  rows?: { cells: { text?: string }[] }[];
}

/** The paragraph styles a creation plan can name. Anything else becomes Normal. */
const CARRIED_STYLES = new Set(['Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Caption', 'Code']);

export interface ReferenceOptions {
  preset?: Preset;
  scenario?: Scenario;
  register?: DesignRegister;
  /** Page geometry, millimetres, passed straight through to the creation plan. */
  page?: { size?: 'A4' | 'Letter'; marginsMm?: { top?: number; right?: number; bottom?: number; left?: number } };
  pageNumberStyle?: 'page' | 'pageOfTotal';
  header?: string;
  footer?: string;
  pageNumbers?: boolean;
  /** Total width the reference's tables are re-laid out across, in millimetres. */
  tableWidthMm?: number;
  /** Treat each table's first row as a header row (`w:tblHeader`). */
  headerRows?: boolean;
}

/**
 * The parse module's optional presentational observation, indexed by block id.
 *
 * Absent unless the caller asked the parser for it (`parseFormatting`). Without it
 * the plan carries text and structure only — the honest outcome, because a
 * semantic parse does not know how anything looked.
 */
export interface DocumentFormatting {
  paragraphs: {
    blockId: string;
    /**
     * Structural anchor, e.g. `word/document.xml!/w:document/w:body/w:tbl[1]/w:tr[2]/w:tc[3]/w:p[1]`.
     * The only join the parse offers between an observed paragraph and the cell
     * of a plan table, so it is declared here rather than looked up by text.
     */
    pointer?: string;
    alignment: string | null;
    /** Effective first-line indent, style chain already resolved. */
    indent?: { firstLineTwips?: number; firstLineChars?: number };
    runs: { text: string; bold?: boolean; size?: number; eastAsia?: string }[];
  }[];
}

/** The size a paragraph renders at when the observation does not say. */
const FALLBACK_SIZE_PT = 12;

/** A paragraph inside a table cell, addressed by the table/row/cell it sits in. */
const CELL_POINTER = /\/w:tbl\[(\d+)\]\/w:tr\[(\d+)\]\/w:tc\[(\d+)\]\/w:p\[\d+\]$/u;

/** The alignments a creation plan can name, keyed by the OOXML `w:jc` value. */
const ALIGNMENTS: Record<string, 'left' | 'center' | 'right' | 'both'> = {
  left: 'left', center: 'center', right: 'right', both: 'both', start: 'left', end: 'right',
};

export interface ReferenceReport {
  /** Blocks the reference carried, by kind. */
  carried: { paragraphs: number; headings: number; tables: number; lists: number; listItems: number; aligned: number; boldAsHeading: number; sized: number; faced: number; cellFormats: number; indented: number };
  /** Blocks the plan could not express, with the reason. Never silently dropped. */
  skipped: { kind: string; reason: string }[];
  /** Objects a plan-built document cannot carry at all, counted from the parse. */
  notCarriedByDesign: { comments: number; footnotes: number; endnotes: number; revisions: number };
  /** Style ids on the reference that have no plan equivalent, with their counts. */
  unmappedStyles: Record<string, number>;
  notes: string[];
}

const clampLevel = (level: number | undefined): 1 | 2 | 3 => {
  const value = Math.max(1, Math.min(3, Math.round(level ?? 1)));
  return value as 1 | 2 | 3;
};

/**
 * Consecutive items of one numbering definition become one list block. A list is
 * a structure, not a run of paragraphs that happen to start with a marker, and a
 * creation plan has a real list block for it.
 *
 * The parse reports membership (`numId`) but not the marker format, so a list is
 * carried as bulleted unless the reference used a numbered list *style*; a plan
 * cannot name a numbering definition the reference owns.
 */
function listKeyOf(block: SemanticBlock): { ordered: boolean; key: string } | undefined {
  const list = block.list ?? undefined;
  const styleId = block.styleId ?? null;
  const byStyle = styleId === null ? undefined : /^ListNumber/iu.test(styleId) ? 'number' : /^List(Bullet|Paragraph)/iu.test(styleId) ? 'bullet' : undefined;
  if (!list && !byStyle) return undefined;
  const ordered = byStyle === 'number';
  const key = list?.numId !== undefined ? `num:${list.numId}` : `style:${styleId ?? 'list'}`;
  return { ordered, key };
}

export function planFromContent(
  content: SemanticContent,
  options: ReferenceOptions = {},
  formatting?: DocumentFormatting,
): { plan: CreatePlan; report: ReferenceReport } {
  const blocks: Record<string, unknown>[] = [];
  const carried = { paragraphs: 0, headings: 0, tables: 0, lists: 0, listItems: 0, aligned: 0, boldAsHeading: 0, sized: 0, faced: 0, cellFormats: 0, indented: 0 };
  const skipped: { kind: string; reason: string }[] = [];
  const unmappedStyles: Record<string, number> = {};
  const notes: string[] = [];
  const tableWidth = options.tableWidthMm ?? 152;
  let sequence = 0;
  const nextId = () => `b${++sequence}`;

  // Presentational facts, when the caller asked the parser for them. A reference
  // that expresses its hierarchy with direct run formatting has no paragraph
  // styles at all, so this is the only place that hierarchy can come from.
  const formattingById = new Map((formatting?.paragraphs ?? []).map((entry) => [entry.blockId, entry]));
  const largestSize = Math.max(0, ...(formatting?.paragraphs ?? []).flatMap((entry) => entry.runs.map((run) => run.size ?? 0)));
  const alignmentOf = (blockId: string) => {
    const jc = formattingById.get(blockId)?.alignment;
    return jc ? ALIGNMENTS[jc] : undefined;
  };
  /**
   * The first-line indent of one observed paragraph, in points.
   *
   * OOXML can state it in twips or in hundredths of a character, and the
   * character form is what Word honours, so the character form is preferred when
   * it is there — converted with the size the paragraph actually renders at, the
   * same size the rebuilt plan will state on its run.
   */
  const firstLineIndentPt = (entry: { indent?: { firstLineTwips?: number; firstLineChars?: number }; runs: { text: string; size?: number }[] } | undefined, size: number | undefined) => {
    const indent = entry?.indent;
    if (!indent) return undefined;
    if (indent.firstLineChars !== undefined) {
      return Math.round((indent.firstLineChars / 100) * (size ?? FALLBACK_SIZE_PT) * 100) / 100;
    }
    if (indent.firstLineTwips !== undefined) return Math.round((indent.firstLineTwips / 20) * 100) / 100;
    return undefined;
  };

  /** A paragraph is a visual heading when every run that has text is bold. */  const isAllBold = (blockId: string) => {
    const runs = formattingById.get(blockId)?.runs.filter((run) => run.text.trim().length > 0) ?? [];
    return runs.length > 0 && runs.every((run) => run.bold === true);
  };
  const sizeOf = (blockId: string) => Math.max(0, ...(formattingById.get(blockId)?.runs ?? []).map((run) => run.size ?? 0));
  /**
   * The run whose declared face and size stand for the whole paragraph.
   *
   * A plan paragraph is rebuilt as one run, so a reference paragraph whose runs
   * disagree cannot be reproduced exactly. The run carrying the most text wins,
   * and the report says the rest was not carried — an approximation stated, not
   * hidden.
   */
  const dominantRun = (blockId: string) => {
    const runs = (formattingById.get(blockId)?.runs ?? []).filter((run) => run.text.trim().length > 0);
    if (!runs.length) return undefined;
    return runs.reduce((best, run) => (run.text.length > best.text.length ? run : best));
  };
  /**
   * Per-cell formatting, keyed `table → "row:cell"`.
   *
   * The k-th `w:tbl` in document order is the k-th table block in document order,
   * which is the assumption this join rests on; a table nested inside a cell
   * would shift the outer count, and this does not attempt to untangle that. A
   * cell whose paragraphs disagree keeps the format of its longest paragraph.
   */
  const cellFormatsByTable = new Map<number, Map<string, { size?: number; eastAsia?: string; bold?: boolean; text: string }>>();
  for (const entry of formatting?.paragraphs ?? []) {
    const match = CELL_POINTER.exec(entry.pointer ?? '');
    if (!match) continue;
    const runs = entry.runs.filter((run) => run.text.trim().length > 0);
    if (!runs.length) continue;
    const dominant = runs.reduce((best, run) => (run.text.length > best.text.length ? run : best));
    if (dominant.size === undefined && dominant.eastAsia === undefined && dominant.bold === undefined) continue;
    const cellKey = `${match[2]}:${match[3]}`;
    const text = runs.map((run) => run.text).join('').trim();
    const tableIndex = Number(match[1]);
    const cells = cellFormatsByTable.get(tableIndex) ?? new Map();
    cellFormatsByTable.set(tableIndex, cells);
    const previous = cells.get(cellKey);
    if (!previous || text.length > previous.text.length) {
      cells.set(cellKey, {
        ...(dominant.size !== undefined ? { size: dominant.size } : {}),
        ...(dominant.eastAsia ? { eastAsia: dominant.eastAsia } : {}),
        ...(dominant.bold !== undefined ? { bold: dominant.bold } : {}),
        text,
      });
    }
  }

  const source = content.semantic.blocks as SemanticBlock[];
  for (let index = 0; index < source.length; index += 1) {
    const block = source[index] as unknown as Record<string, unknown>;
    const kind = String(block.kind ?? 'unknown');

    if (kind === 'table') {
      const rows = (block.rows as { cells: { text?: string }[] }[] | undefined) ?? [];
      const columnCount = Number(block.gridColumns ?? rows[0]?.cells.length ?? 0);
      if (!rows.length || columnCount === 0) {
        skipped.push({ kind: 'table', reason: 'the reference table has no rows or no columns' });
        continue;
      }
      const tableIndex = carried.tables + 1;
      const observed = cellFormatsByTable.get(tableIndex);
      let cellFormats: (Record<string, unknown> | null)[][] | undefined;
      if (observed) {
        cellFormats = rows.map((row, rowIndex) => row.cells.map((_cell, cellIndex) => {
          const cell = observed.get(`${rowIndex + 1}:${cellIndex + 1}`);
          if (!cell) return null;
          carried.cellFormats += 1;
          return {
            ...(cell.size !== undefined ? { size: cell.size } : {}),
            ...(cell.eastAsia ? { eastAsia: cell.eastAsia } : {}),
            ...(cell.bold !== undefined ? { bold: cell.bold } : {}),
          };
        }));
        if (!cellFormats.some((row) => row.some((cell) => cell !== null))) cellFormats = undefined;
      }
      blocks.push({
        kind: 'table', id: nextId(),
        rows: rows.map((row) => row.cells.map((cell) => cell.text ?? '')),
        // A reference's own column widths are geometry the creation plan
        // rebuilds, so the columns are re-laid out across the printable width.
        columnWidthsMm: Array.from({ length: columnCount }, () => Math.round((tableWidth / columnCount) * 10) / 10),
        ...(options.headerRows ? { header: true } : {}),
        ...(cellFormats ? { cellFormats } : {}),
      });
      carried.tables += 1;
      continue;
    }

    if (kind !== 'paragraph' && kind !== 'heading') {
      skipped.push({ kind, reason: 'a creation plan has no block for this node kind' });
      continue;
    }

    const styleId = (block.styleId as string | null) ?? null;
    const listKey = listKeyOf(block as unknown as SemanticBlock);

    // Gather a whole run of items from one numbering definition into one list.
    if (listKey) {
      const items: string[] = [];
      let cursor = index;
      while (cursor < source.length) {
        const candidate = source[cursor] as SemanticBlock;
        if (candidate.kind !== 'paragraph' || listKeyOf(candidate)?.key !== listKey.key) break;
        items.push(String(candidate.text ?? ''));
        cursor += 1;
      }
      index = cursor - 1;
      if (items.length) {
        blocks.push({ kind: 'list', id: nextId(), items, ...(listKey.ordered ? { ordered: true } : {}) });
        carried.lists += 1;
        carried.listItems += items.length;
      }
      continue;
    }

    const text = String(block.text ?? '');
    const blockId = String(block.id ?? '');
    let style: string;
    if (kind === 'heading') {
      style = `Heading${clampLevel(block.level as number | undefined)}`;
      carried.headings += 1;
    } else if (styleId && CARRIED_STYLES.has(styleId)) {
      style = styleId;
      carried.paragraphs += 1;
    } else {
      style = 'Normal';
      carried.paragraphs += 1;
      if (styleId) unmappedStyles[styleId] = (unmappedStyles[styleId] ?? 0) + 1;
    }

    // A reference that marks its headings with direct run formatting has no
    // paragraph style to carry, so an entirely bold paragraph is the heading.
    // The largest bold run size in the document identifies the title.
    if (kind !== 'heading' && text.trim() !== '' && isAllBold(blockId)) {
      const isLargest = sizeOf(blockId) > 0 && sizeOf(blockId) === largestSize;
      style = isLargest ? 'Title' : 'Heading1';
      carried.boldAsHeading += 1;
    }

    // Alignment is an observed fact with a direct plan field, so it is carried
    // as-is rather than inferred from the paragraph's role.
    const alignment = alignmentOf(blockId);
    if (alignment) carried.aligned += 1;

    // Run-level size and face are carried too: a reference that states them
    // directly is stating a typographic fact, and without them the rebuilt
    // document would fall back to the preset's sizes — which is how a 10.5pt
    // `编号` line and a 14pt appendix heading used to come back at body size.
    // Colour is deliberately not carried: a creation plan's palette is the
    // register's, and a black-on-white formal record has nothing to preserve.
    const face = dominantRun(blockId);
    const run: Record<string, unknown> = { text };
    if (face?.size !== undefined && face.size > 0) {
      run.size = face.size;
      carried.sized += 1;
    }
    if (face?.eastAsia) {
      run.eastAsia = face.eastAsia;
      carried.faced += 1;
    }

    // The first-line indent is carried as a number of points, because a plan
    // states lengths in points; the observation says it in the two units OOXML
    // uses. `0` is carried like any other value: it is how a signature or
    // address line stops inheriting the register's two-character indent.
    const indentPt = firstLineIndentPt(formattingById.get(blockId), face?.size);
    if (indentPt !== undefined) carried.indented += 1;

    blocks.push({
      kind: 'paragraph', id: nextId(), style, runs: [run],
      ...(alignment ? { alignment } : {}),
      ...(indentPt !== undefined ? { firstLineIndentPt: indentPt } : {}),
    });
  }

  const annotations = content.semantic.annotations ?? [];
  const countOf = (kind: string) => annotations.filter((annotation: { kind: string }) => annotation.kind === kind).length;

  const document: Record<string, unknown> = { blocks };
  if (options.preset) document.preset = options.preset;
  if (options.scenario) document.scenario = options.scenario;
  if (options.register) document.register = options.register;
  if (options.page) document.page = options.page;
  if (options.pageNumberStyle) document.pageNumberStyle = options.pageNumberStyle;
  if (options.header !== undefined) document.header = options.header;
  if (options.footer !== undefined) document.footer = options.footer;
  if (options.pageNumbers !== undefined) document.pageNumbers = options.pageNumbers;

  if (carried.tables && !options.headerRows) {
    notes.push('Table header rows were not marked: a parsed table does not state whether its first row is a heading, so pass headerRows to assert it.');
  }
  if (carried.lists) {
    notes.push('Lists are carried as bulleted unless the reference used a numbered list style: the parse reports list membership and grouping, not the marker format.');
  }
  if (Object.keys(unmappedStyles).length) {
    notes.push('Styles with no creation-plan equivalent were rendered as Normal; their look is the preset\'s, not the reference\'s.');
  }
  if (!formatting) {
    notes.push('No presentational observation was requested, so paragraph alignment and run formatting were not carried: a semantic parse does not know how anything looked.');
  } else {
    notes.push(`Presentational facts read: ${carried.aligned} paragraph(s) kept their alignment, ${carried.boldAsHeading} entirely-bold paragraph(s) became headings, ${carried.sized} run size(s) and ${carried.faced} East Asian face(s) were carried.`);
    if (carried.indented) notes.push(`${carried.indented} paragraph(s) kept their own first-line indent — including the ones whose indent is zero, which is how a signature or address line stops inheriting the register's.`);
    if (carried.cellFormats) notes.push(`${carried.cellFormats} table cell(s) kept their own size and face; the rest of a cell's look is the register's table scheme.`);
  }
  if (carried.sized) {
    notes.push('Run sizes are carried per paragraph, not per run: a paragraph whose runs declared different sizes kept the size of its dominant run.');
  }
  notes.push('Colours still come from the preset: a creation plan carries a style, an alignment and a run size and face, not a run palette.');

  return {
    plan: { kind: 'create', document } as CreatePlan,
    report: {
      carried,
      skipped,
      notCarriedByDesign: {
        comments: countOf('comment'),
        footnotes: countOf('footnote'),
        endnotes: countOf('endnote'),
        revisions: annotations.filter((annotation: { kind: string }) => String(annotation.kind).includes('revision')).length,
      },
      unmappedStyles,
      notes,
    },
  };
}

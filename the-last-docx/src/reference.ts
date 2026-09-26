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

export interface ReferenceReport {
  /** Blocks the reference carried, by kind. */
  carried: { paragraphs: number; headings: number; tables: number; lists: number; listItems: number };
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

export function planFromContent(content: SemanticContent, options: ReferenceOptions = {}): { plan: CreatePlan; report: ReferenceReport } {
  const blocks: Record<string, unknown>[] = [];
  const carried = { paragraphs: 0, headings: 0, tables: 0, lists: 0, listItems: 0 };
  const skipped: { kind: string; reason: string }[] = [];
  const unmappedStyles: Record<string, number> = {};
  const notes: string[] = [];
  const tableWidth = options.tableWidthMm ?? 152;
  let sequence = 0;
  const nextId = () => `b${++sequence}`;

  const source = content.semantic.blocks as SemanticBlock[];
  for (let index = 0; index < source.length; index += 1) {
    const block = source[index] as unknown as Record<string, unknown>;    const kind = String(block.kind ?? 'unknown');

    if (kind === 'table') {
      const rows = (block.rows as { cells: { text?: string }[] }[] | undefined) ?? [];
      const columnCount = Number(block.gridColumns ?? rows[0]?.cells.length ?? 0);
      if (!rows.length || columnCount === 0) {
        skipped.push({ kind: 'table', reason: 'the reference table has no rows or no columns' });
        continue;
      }
      blocks.push({
        kind: 'table', id: nextId(),
        rows: rows.map((row) => row.cells.map((cell) => cell.text ?? '')),
        // A reference's own column widths are geometry the creation plan
        // rebuilds, so the columns are re-laid out across the printable width.
        columnWidthsMm: Array.from({ length: columnCount }, () => Math.round((tableWidth / columnCount) * 10) / 10),
        ...(options.headerRows ? { header: true } : {}),
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
    // Runs are rebuilt as plain text: a creation plan carries text and a style,
    // not the reference's run-level character formatting.
    blocks.push({ kind: 'paragraph', id: nextId(), style, runs: [{ text }] });
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
  notes.push('Run-level character formatting is not carried: the preset decides faces, sizes and colours.');

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

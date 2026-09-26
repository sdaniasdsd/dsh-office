import type { NodeAnchor, SemanticId } from '@dsh-office-profile/docx-parse';

/** A target obtained from docx-parse; the anchor stays owned by that module's contract. */
export interface DocxEditTarget {
  semanticId: SemanticId;
  anchor: NodeAnchor;
}

export type RevisionMode = 'respectDocument' | 'track' | 'untracked';

/** Office JS `Word.Alignment`; the four values the engine reads back. */
export type ParagraphAlignment = 'Left' | 'Centered' | 'Right' | 'Justified';

export interface ReplaceTextEdit {
  kind: 'replaceText';
  target: DocxEditTarget;
  find: string;
  replace: string;
  /** Zero-based when the text occurs more than once; omitted means it must be unique. */
  occurrence?: number;
  matchCase?: boolean;
  revision?: RevisionMode;
}

export interface AddCommentEdit {
  kind: 'addComment';
  target: DocxEditTarget;
  /** Omit to anchor the comment to the whole paragraph. */
  quote?: string;
  occurrence?: number;
  text: string;
  author: { name: string; initials?: string; email?: string };
}

export interface ResolveRevisionsEdit {
  kind: 'resolveRevisions';
  target: DocxEditTarget;
  decision: 'accept' | 'reject';
}

/**
 * Direct run-level character formatting applied to every run of the target
 * paragraph. Kept separate from `styleId` on purpose: assigning a paragraph
 * style only changes what a reader sees when the document already defines that
 * style, whereas direct formatting is what makes a hierarchy visible in a
 * document whose `styles.xml` carries no heading styles.
 */
export interface ParagraphFontFormat {
  bold?: boolean;
  italic?: boolean;
  /** Points, as Office JS. */
  size?: number;
  /** '#RRGGBB', the same shape the engine reads back. */
  color?: string;
  /** Font family name. */
  name?: string;
}

/**
 * In-place paragraph formatting.
 *
 * This writes only the target paragraph's `w:pPr` (and, when `font` is given,
 * the direct run properties of its runs). Every part the edit does not touch —
 * `comments.xml`, `footnotes.xml`, headers, numbering, custom XML — keeps its
 * original bytes, which is why this is the operation to reach for when a
 * document has to be re-styled without losing its annotations. The rebuilding
 * modules cannot make that promise: a new package built from a plan has no
 * comments or footnotes to carry over.
 *
 * At least one formatting property is required; a request that would change
 * nothing fails instead of reporting a silent no-op.
 */
export interface FormatParagraphEdit {
  kind: 'formatParagraph';
  target: DocxEditTarget;
  /** Paragraph style id (`w:pStyle`). The style must already exist in the document. */
  styleId?: string;
  /**
   * Outline level as Office JS reports it: 1-9 for heading levels, 10 for body
   * text. This is NOT the zero-based `w:outlineLvl` stored in the XML; it is
   * the engine's own read/write convention, so what you set is what reads back.
   */
  outlineLevel?: number;
  alignment?: ParagraphAlignment;
  /** Spacing before the paragraph, in points. */
  spaceBefore?: number;
  /** Spacing after the paragraph, in points. */
  spaceAfter?: number;
  /** Line spacing, in points. */
  lineSpacing?: number;
  font?: ParagraphFontFormat;
}

export type DocxEditOperation = ReplaceTextEdit | AddCommentEdit | ResolveRevisionsEdit | FormatParagraphEdit | FormatTableEdit | InsertParagraphEdit;

export type TableCellVerticalAlignment = 'top' | 'center' | 'bottom';
export type TableAlignment = 'Left' | 'Center' | 'Right';
export type TableBorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'none';
export type TableBorderEdge = 'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV';

export interface TableBorderEdgeSpec {
  style?: TableBorderStyle;
  /** Points; the format stores eighths of a point. */
  size?: number;
  /** `'#RRGGBB'`, `'auto'`, or a theme colour name. */
  color?: string;
}

/**
 * A table's own geometry - not a reusable style.
 *
 * A table style only says what a table *may* look like; what decides the layout
 * is the table element itself. A `<w:tbl>` whose `<w:tblGrid>` has bare
 * `<w:gridCol/>` children carries **no preferred widths**, so a renderer falls
 * back to the minimum width the content needs and a cell reading `A1` wraps to
 * two lines. No style definition can repair that, which is why applying a table
 * style and giving the table its geometry are separate steps, and why this edit
 * lives beside the content edits rather than in the style module.
 *
 * Mature libraries agree: python-docx puts `autofit`/`column.width`/`cell.width`
 * on the table object, Apache POI puts them on `XWPFTable`, docx4j on the `Tbl`.
 *
 * Units are points, as everywhere else here.
 */
export interface TableGeometry {
  /** `w:tblStyle`: the named table style this table uses. */
  styleId?: string;
  /** Total width (`w:tblW`). */
  width?: number;
  /** `fixed` makes the columns honour the grid, which is what stops the wrapping. */
  layout?: 'fixed' | 'autofit';
  alignment?: TableAlignment;
  /** Per-column widths; must total `width` when both are given. */
  columnWidths?: number[];
  cellMargins?: { top?: number; left?: number; bottom?: number; right?: number };
  borders?: Partial<Record<TableBorderEdge, TableBorderEdgeSpec>>;
  /** Repeat the first row as a header across page breaks (`w:tblHeader`). */
  headerRow?: boolean;
  cellVerticalAlignment?: TableCellVerticalAlignment;
}

export interface FormatTableEdit extends TableGeometry {
  kind: 'formatTable';
  target: DocxEditTarget;
}

/**
 * Insert a new paragraph next to a parsed target.
 *
 * Until this kind existed the module could change what a paragraph said but not
 * add one, so a request like "add a sentence after this paragraph" - a routine
 * ask in any review workflow, and supported by python-docx's `add_paragraph`,
 * POI's `insertNewParagraph` and docx4j's own `Paragraph.insertParagraph` - had
 * no expression at all.
 */
export interface InsertParagraphEdit {
  kind: 'insertParagraph';
  target: DocxEditTarget;
  text: string;
  position: 'Before' | 'After';
  /** `respectDocument` keeps the document's tracking mode, as `replaceText` does. */
  revision?: RevisionMode;
  /** Style id for the new paragraph. Omit to inherit the target's properties. */
  styleId?: string;
}

export interface EditPlan {
  edits: DocxEditOperation[];
  author?: { name: string; initials?: string; email?: string };
}

export interface AppliedEdit {
  kind: DocxEditOperation['kind'];
  semanticId: SemanticId;
  changed: number;
  /** Only structural facts are returned; document text is not copied into telemetry. */
  details?: Record<string, string | number | boolean>;
}

/**
 * What a `formatParagraph` edit asked for, kept so `verify()` can read the
 * saved package back and assert the formatting is actually there. Recording an
 * intent is not evidence that it landed.
 */
export interface ExpectedParagraphFormat {
  /** The paragraph's text, used to find it again in the saved output. */
  text: string;
  semanticId: SemanticId;
  styleId?: string;
  outlineLevel?: number;
  alignment?: ParagraphAlignment;
  spaceBefore?: number;
  spaceAfter?: number;
  lineSpacing?: number;
  bold?: boolean;
  italic?: boolean;
  size?: number;
  color?: string;
  name?: string;
}

/**
 * What a `formatTable` edit asked for. Verification re-reads the saved table and
 * asserts each property is present, because writing geometry and having it
 * survive the save are different facts.
 */
export interface ExpectedTableFormat {
  /** Zero-based index of the table in document order. */
  tableIndex: number;
  columnCount: number;
  applied: string[];
  styleId?: string;
  widthPt?: number;
  layout?: 'fixed' | 'autofit';
  columnWidthsPt?: number[];
  headerRow?: boolean;
}

/**
 * What an `insertParagraph` edit asked for. `tracked` is recorded because the
 * difference between "the sentence is there" and "the sentence is there as a
 * Word revision" is the whole point of the request.
 */
export interface ExpectedInsertion {
  text: string;
  tracked: boolean;
}

export interface EngineEditResult {
  bytes: Uint8Array;
  edits: AppliedEdit[];
  touchedParts: string[];
  paragraphCount: number;
  tableCount: number;
  expectedParagraphTexts: string[];
  expectedCommentTexts: string[];
  expectedFormats: ExpectedParagraphFormat[];
  expectedTableFormats: ExpectedTableFormat[];
  expectedInsertions: ExpectedInsertion[];
}

export interface PackageInspection {
  paragraphCount: number;
  tableCount: number;
  commentCount: number;
  trackedChangeCount: number;
  byteLength: number;
}

/**
 * Named-style authoring for `styles.xml`.
 *
 * This module owns one OOXML part — the style definitions — the way docx4j's
 * `StyleDefinitionsPart`, Apache POI's `XWPFStyles` and the Open XML SDK's
 * `StyleDefinitionsPart` each own theirs. It does not touch document content:
 * applying a style to a paragraph stays with the editing module, so the two
 * concerns do not both grow a copy of the other.
 *
 * **Units are points**, as in Office JS and in `docx-edit`'s `formatParagraph`.
 * The engine converts to what the format actually stores - twips (`pt * 20`),
 * half-points (`pt * 2`), eighths of a point for border sizes (`pt * 8`) - so a
 * caller never does unit arithmetic and cannot silently write the wrong scale.
 */

export type StyleType = 'paragraph' | 'character' | 'table';
export type StyleAlignment = 'Left' | 'Centered' | 'Right' | 'Justified';
export type LineRule = 'auto' | 'exact' | 'atLeast';
export type BorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'none';

/** Points. `line`'s meaning depends on `lineRule`, as in OOXML. */
export interface StyleSpacing {
  before?: number;
  after?: number;
  line?: number;
  lineRule?: LineRule;
}

/** Points. `hanging` and `firstLine` are mutually exclusive in OOXML. */
export interface StyleIndentation {
  left?: number;
  right?: number;
  firstLine?: number;
  hanging?: number;
}

export interface StyleBorderEdge {
  style?: BorderStyle;
  /** Points; the format stores eighths of a point. */
  size?: number;
  /** `'#RRGGBB'`, `'auto'`, or a theme colour name. */
  color?: string;
  /** Points of space between the border and the text. */
  space?: number;
}

export type TableBorderEdge = 'top' | 'left' | 'bottom' | 'right' | 'insideH' | 'insideV';

export interface StyleParagraphFormat {
  /**
   * Zero-based outline level (`w:outlineLvl`), 0-8. This is the raw OOXML
   * value, not Office JS's 1-based `outlineLevel`: a style definition is the
   * one place where the stored number is the honest one to write.
   */
  outlineLevel?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  alignment?: StyleAlignment;
  spacing?: StyleSpacing;
  indentation?: StyleIndentation;
  bottomBorder?: StyleBorderEdge;
}

export interface StyleFontFormat {
  bold?: boolean;
  italic?: boolean;
  /** Points; the format stores half-points. */
  size?: number;
  /** `'#RRGGBB'` or `'auto'`. */
  color?: string;
  font?: { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string };
}

export interface StyleTableFormat {
  /** Points; the format stores twips. */
  width?: number;
  /** Points of indent before the table (`w:tblInd`). */
  indent?: number;
  layout?: 'fixed' | 'autofit';
  /** Points. */
  cellMargins?: { top?: number; left?: number; bottom?: number; right?: number };
  borders?: Partial<Record<TableBorderEdge, StyleBorderEdge>>;
}

/**
 * One `w:style` definition. `styleId` is what content references in
 * `w:pStyle`/`w:rStyle`/`w:tblStyle`; `name` is the display name Word shows.
 */
export interface StyleSpec {
  styleId: string;
  name: string;
  type?: StyleType;
  basedOn?: string;
  next?: string;
  link?: string;
  /** `w:default="1"`: the style applied when content names none. */
  default?: boolean;
  quickFormat?: boolean;
  paragraph?: StyleParagraphFormat;
  run?: StyleFontFormat;
  table?: StyleTableFormat;
}

export interface StylePlan {
  styles: StyleSpec[];
}

export type StyleAction = 'created' | 'replaced' | 'unchanged';

export interface DefinedStyle {
  styleId: string;
  name: string;
  type: StyleType;
  action: StyleAction;
  /** The property groups the definition carries, for telemetry and review. */
  properties: string[];
}

export interface StyleSummary {
  styleId: string;
  name: string;
  type: StyleType;
  basedOn?: string;
  default: boolean;
  /** The `w:pPr`/`w:rPr`/`w:tblPr` children the definition actually carries. */
  carries: string[];
}

/**
 * What the document's own style definitions say, and which references they fail
 * to satisfy.
 *
 * The three reference buckets are kept apart on purpose. Word synthesizes the
 * built-in styles it knows even when a document does not define them, so a
 * reference to `Heading2` in a document with no `Heading2` definition is not
 * broken - it is *renderer-dependent*, which is a different and milder fact
 * than a reference to a custom style that nothing defines.
 */
export interface StyleInventory {
  hasStylesPart: boolean;
  styles: StyleSummary[];
  defaultParagraphStyle?: string;
  defaultTableStyle?: string;
  references: {
    defined: string[];
    builtInUndefined: string[];
    unknownUndefined: string[];
  };
}

export interface EngineStyleResult {
  bytes: Uint8Array;
  defined: DefinedStyle[];
  inventory: StyleInventory;
}

/** A property the caller asked for, kept so verification can read it back. */
export interface ExpectedStyleDefinition {
  styleId: string;
  name: string;
  type: StyleType;
  /** Property groups requested, e.g. `paragraph.spacing`, `run.bold`. */
  properties: string[];
}

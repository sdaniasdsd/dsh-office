import type { DesignTokens } from './design';
import { NS, xml } from './xml';

/** The `w:numId` a bullet list references. */
export const BULLET_NUM_ID = 1;
/** The `w:numId` an ordered list references. */
export const DECIMAL_NUM_ID = 2;

/** The hanging indent a list level uses when its register does not say otherwise, in twips. */
export const DEFAULT_LIST_INDENT = { left: 720, hanging: 360 } as const;

/**
 * The numbering definitions a document with list blocks needs.
 *
 * `createPackage` already ships an empty numbering part; this fills it only
 * when a list block is present, so a document without lists carries no invented
 * numbering for a reader to reconcile.
 *
 * The bullet is the literal U+2022 with no font override, on purpose. Word's own
 * definitions point at the Symbol font's private-use `F0B7`, which any renderer
 * that maps the plain character instead draws as a missing glyph, and a list
 * marker is not worth a renderer-specific trap.
 */
export function numberingXml(design: DesignTokens): string {
  const { left, hanging } = design.list?.indent ?? DEFAULT_LIST_INDENT;
  const level = (format: 'bullet' | 'decimal', text: string) =>
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${xml(text)}"/>`
    + `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr></w:lvl>`;
  return `<w:numbering ${NS}>`
    + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${level('bullet', '\u2022')}</w:abstractNum>`
    + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${level('decimal', '%1.')}</w:abstractNum>`
    + `<w:num w:numId="${BULLET_NUM_ID}"><w:abstractNumId w:val="0"/></w:num>`
    + `<w:num w:numId="${DECIMAL_NUM_ID}"><w:abstractNumId w:val="1"/></w:num>`
    + `</w:numbering>`;
}

/** The `w:numPr` a list item carries. `w:pPr`'s `w:numPr` follows `w:pStyle`. */
export function numPrXml(ordered: boolean): string {
  return `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${ordered ? DECIMAL_NUM_ID : BULLET_NUM_ID}"/></w:numPr>`;
}

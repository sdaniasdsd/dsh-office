import type { StyleBorderEdge, StyleSpec, TableBorderEdge } from '../domain/docx-styles';
import { colorValue, toEighths, toHalfPoints, toPointMeasure, toTwips, xml } from './xml';

/**
 * Child sequences declared by ECMA-376 for the containers this module writes.
 * The emitter builds each container through these tables, so the order is right
 * by construction rather than by remembering; the verifier re-checks it on the
 * finished part.
 */
export const STYLE_CHILD_ORDER = [
  'name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority',
  'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose',
  'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr',
] as const;

export const PPR_CHILD_ORDER = [
  'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
  'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
  'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
  'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
] as const;

export const RPR_CHILD_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
  'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
  'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
  'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
  'specVanish', 'oMath',
] as const;

export const TBLPR_CHILD_ORDER = [
  'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
  'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd',
  'tblLayout', 'tblCellMar', 'tblLook',
] as const;

export const TBLBORDERS_CHILD_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;

/** Emit `entries` in `order`, skipping the ones that are undefined. */
function assemble(entries: Record<string, string | undefined>, order: readonly string[]): string {
  return order.map((name) => entries[name] ?? '').join('');
}

function borderEdge(tag: string, edge: StyleBorderEdge): string {
  const attributes: string[] = [];
  if (edge.style !== undefined) attributes.push(`w:val="${xml(edge.style)}"`);
  attributes.push(`w:sz="${toEighths(edge.size ?? 0.5)}"`);
  attributes.push(`w:space="${toPointMeasure(edge.space ?? 0)}"`);
  attributes.push(`w:color="${xml(colorValue(edge.color ?? 'auto'))}"`);
  return `<w:${tag} ${attributes.join(' ')}/>`;
}

/**
 * OOXML's on/off elements: the bare element means true, and an explicit
 * `w:val="0"` is what states false. Omitting the element leaves the value to
 * whatever the style inherits, which is not the same as stating false.
 */
function onOffTag(name: string, value: boolean | undefined): string | undefined {
  if (value === undefined) return undefined;
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
}

function paragraphProperties(spec: StyleSpec): string | undefined {
  const paragraph = spec.paragraph;
  if (!paragraph) return undefined;
  const spacing = paragraph.spacing;
  const indentation = paragraph.indentation;
  const entries: Record<string, string | undefined> = {
    keepNext: onOffTag('keepNext', paragraph.keepNext),
    keepLines: onOffTag('keepLines', paragraph.keepLines),
    pBdr: paragraph.bottomBorder ? `<w:pBdr>${borderEdge('bottom', paragraph.bottomBorder)}</w:pBdr>` : undefined,
    spacing: spacing
      ? `<w:spacing${spacing.before !== undefined ? ` w:before="${toTwips(spacing.before)}"` : ''}${spacing.after !== undefined ? ` w:after="${toTwips(spacing.after)}"` : ''}${spacing.line !== undefined ? ` w:line="${toTwips(spacing.line)}"` : ''}${spacing.lineRule !== undefined ? ` w:lineRule="${xml(spacing.lineRule)}"` : ''}/>`
      : undefined,
    ind: indentation
      ? `<w:ind${indentation.left !== undefined ? ` w:left="${toTwips(indentation.left)}"` : ''}${indentation.right !== undefined ? ` w:right="${toTwips(indentation.right)}"` : ''}${indentation.firstLine !== undefined ? ` w:firstLine="${toTwips(indentation.firstLine)}"` : ''}${indentation.hanging !== undefined ? ` w:hanging="${toTwips(indentation.hanging)}"` : ''}/>`
      : undefined,
    jc: paragraph.alignment ? `<w:jc w:val="${xml(paragraph.alignment)}"/>` : undefined,
    outlineLvl: paragraph.outlineLevel !== undefined ? `<w:outlineLvl w:val="${paragraph.outlineLevel}"/>` : undefined,
  };
  const inner = assemble(entries, PPR_CHILD_ORDER);
  return inner ? `<w:pPr>${inner}</w:pPr>` : undefined;
}

function runProperties(spec: StyleSpec): string | undefined {
  const run = spec.run;
  if (!run) return undefined;
  const font = run.font;
  const entries: Record<string, string | undefined> = {
    rFonts: font
      ? `<w:rFonts${font.ascii ? ` w:ascii="${xml(font.ascii)}"` : ''}${font.hAnsi ? ` w:hAnsi="${xml(font.hAnsi)}"` : ''}${font.eastAsia ? ` w:eastAsia="${xml(font.eastAsia)}"` : ''}${font.cs ? ` w:cs="${xml(font.cs)}"` : ''}/>`
      : undefined,
    b: run.bold ? '<w:b/>' : undefined,
    bCs: run.bold ? '<w:bCs/>' : undefined,
    i: run.italic ? '<w:i/>' : undefined,
    iCs: run.italic ? '<w:iCs/>' : undefined,
    color: run.color !== undefined ? `<w:color w:val="${xml(colorValue(run.color))}"/>` : undefined,
    sz: run.size !== undefined ? `<w:sz w:val="${toHalfPoints(run.size)}"/>` : undefined,
    szCs: run.size !== undefined ? `<w:szCs w:val="${toHalfPoints(run.size)}"/>` : undefined,
  };
  const inner = assemble(entries, RPR_CHILD_ORDER);
  return inner ? `<w:rPr>${inner}</w:rPr>` : undefined;
}

function tableProperties(spec: StyleSpec): string | undefined {
  const table = spec.table;
  if (!table) return undefined;
  const borders = table.borders;
  const borderEntries: Record<string, string | undefined> = {};
  if (borders) {
    for (const edge of TBLBORDERS_CHILD_ORDER) {
      const value = borders[edge];
      if (value) borderEntries[edge] = borderEdge(edge, value);
    }
  }
  const borderInner = borders ? assemble(borderEntries, TBLBORDERS_CHILD_ORDER) : '';
  const margins = table.cellMargins;
  const entries: Record<string, string | undefined> = {
    tblW: table.width !== undefined ? `<w:tblW w:w="${toTwips(table.width)}" w:type="dxa"/>` : undefined,
    tblInd: table.indent !== undefined ? `<w:tblInd w:w="${toTwips(table.indent)}" w:type="dxa"/>` : undefined,
    tblBorders: borderInner ? `<w:tblBorders>${borderInner}</w:tblBorders>` : undefined,
    tblLayout: table.layout ? `<w:tblLayout w:type="${xml(table.layout)}"/>` : undefined,
    tblCellMar: margins
      ? `<w:tblCellMar>${(['top', 'left', 'bottom', 'right'] as const)
        .filter((edge) => margins[edge] !== undefined)
        .map((edge) => `<w:${edge} w:w="${toTwips(margins[edge] ?? 0)}" w:type="dxa"/>`)
        .join('')}</w:tblCellMar>`
      : undefined,
  };
  const inner = assemble(entries, TBLPR_CHILD_ORDER);
  return inner ? `<w:tblPr>${inner}</w:tblPr>` : undefined;
}

/** One `w:style` element, with its children in the sequence ECMA-376 declares. */
export function styleElement(spec: StyleSpec): string {
  const type = spec.type ?? 'paragraph';
  const entries: Record<string, string | undefined> = {
    name: `<w:name w:val="${xml(spec.name)}"/>`,
    basedOn: spec.basedOn ? `<w:basedOn w:val="${xml(spec.basedOn)}"/>` : undefined,
    next: spec.next ? `<w:next w:val="${xml(spec.next)}"/>` : undefined,
    link: spec.link ? `<w:link w:val="${xml(spec.link)}"/>` : undefined,
    qFormat: spec.quickFormat ? '<w:qFormat/>' : undefined,
    pPr: paragraphProperties(spec),
    rPr: runProperties(spec),
    tblPr: tableProperties(spec),
  };
  const attributes = [`w:type="${xml(type)}"`, `w:styleId="${xml(spec.styleId)}"`];
  if (spec.default) attributes.push('w:default="1"');
  return `<w:style ${attributes.join(' ')}>${assemble(entries, STYLE_CHILD_ORDER)}</w:style>`;
}

/** The `w:pBdr` edges this module can express, for callers that build one. */
export type BorderEdgeName = TableBorderEdge;

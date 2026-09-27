import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Element, Node } from '@xmldom/xmldom';
import { DocxStylesError } from '../errors';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS = `xmlns:w="${W}" xmlns:r="${R}"`;

export function xml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/** Points to twentieths of a point, the unit `w:spacing`, `w:ind` and `w:tblW` store. */
export const toTwips = (points: number): number => Math.round(points * 20);
/** Points to half-points, the unit `w:sz` stores. */
export const toHalfPoints = (points: number): number => Math.round(points * 2);
/** Points to eighths of a point, the unit a border's `w:sz` stores. */
export const toEighths = (points: number): number => Math.round(points * 8);
/**
 * A border's `w:space` is ST_PointMeasure — whole points, NOT twentieths. Word
 * writes `<w:bottom w:val="single" w:sz="4" w:space="4" .../>` for a half-point
 * rule offset four points from the text; carrying that through `toTwips` would
 * ask for a twentieth of a point and silently place the rule against the glyphs.
 */
export const toPointMeasure = (points: number): number => Math.round(points);

/** `'#RRGGBB'` becomes `'RRGGBB'`; `'auto'` and theme colour names pass through. */
export function colorValue(value: string): string {
  return /^#[0-9a-fA-F]{6}$/.test(value) ? value.slice(1).toUpperCase() : value;
}

export function parseXml(source: string) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) {
    throw new DocxStylesError('SAFETY_POLICY_DENIED', 'DTD and entity declarations are forbidden.');
  }
  try {
    return new DOMParser({ onError: () => { throw new Error('Malformed XML'); } }).parseFromString(source, 'application/xml');
  } catch {
    throw new DocxStylesError('FORMAT_MISMATCH', 'A package XML part is malformed.');
  }
}

export function serializeXml(document: ReturnType<typeof parseXml>): string {
  return new XMLSerializer().serializeToString(document);
}

/** Immediate element children, in document order. */
export function childElements(element: Element): Element[] {
  const out: Element[] = [];
  const nodes = element.childNodes;
  for (let index = 0; index < nodes.length; index += 1) {
    const node: Node | undefined = nodes.item(index) ?? undefined;
    if (node && node.nodeType === 1) out.push(node as Element);
  }
  return out;
}

/** The local name without its namespace prefix. */
export function localNameOf(element: Element): string {
  return element.localName ?? element.nodeName.replace(/^.*:/, '');
}

/**
 * Report children that appear out of the sequence OOXML declares.
 *
 * Element order in these parts is not cosmetic: Word rejects a package whose
 * `w:pPr` puts `w:pBdr` after `w:spacing`, and offers to "repair" the file. The
 * emitter already orders what it writes; this catches a part that arrived out of
 * order from somewhere else, or a hand-written merge that slipped.
 *
 * Elements not named in `order` are ignored, so a schema extension this module
 * does not model cannot produce a false report.
 */
export function orderProblems(element: Element, order: readonly string[], label: string): string[] {
  const rank = new Map(order.map((name, index) => [name, index]));
  const problems: string[] = [];
  let highest = -1;
  let highestName = '';
  for (const child of childElements(element)) {
    const name = localNameOf(child);
    const position = rank.get(name);
    if (position === undefined) continue;
    if (position < highest) problems.push(`${label}: <${name}> appears after <${highestName}>`);
    else { highest = position; highestName = name; }
  }
  return problems;
}

/** Every `w:<attribute>` value of the elements matching `name` anywhere in the tree. */
export function attributeValues(root: Element, name: string, attribute: string): string[] {
  const out: string[] = [];
  const walk = (element: Element): void => {
    for (const child of childElements(element)) {
      if (localNameOf(child) === name) {
        const value = child.getAttribute(attribute) ?? child.getAttributeNS(W, attribute.slice(2));
        if (value) out.push(value);
      }
      walk(child);
    }
  };
  walk(root);
  return out;
}

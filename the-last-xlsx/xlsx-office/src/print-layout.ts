import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface PrintLayoutChange {
  sheet: string;
  orientation?: 'portrait' | 'landscape';
  fitToWidth?: number;
  fitToHeight?: number;
}

function parseXml(source: string, part: string) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error(`Unsafe XML declarations in ${part}.`);
  let parserError = false;
  const document = new DOMParser({ onError: () => { parserError = true; } }).parseFromString(source, 'application/xml');
  if (parserError || !document.documentElement) throw new Error(`Malformed XML in ${part}.`);
  return document;
}

function getElementAttribute(xml: string, elementName: string, attributeName: string): string | undefined {
  const tag = new RegExp(`<(?:(?:[\\w.-]+):)?${elementName}\\b([^>]*)\\/?\\s*>`, 'i').exec(xml)?.[1];
  if (tag === undefined) return undefined;
  const attributes = new RegExp(`(?:^|\\s)${attributeName}\\s*=\\s*(["'])(.*?)\\1`, 'i').exec(tag);
  return attributes?.[2];
}

function replaceElement(xml: string, elementName: string, wanted: Record<string, string | undefined>, insertBefore: RegExp): string {
  const pattern = new RegExp(`<((?:[\\w.-]+:)?${elementName})\\b([^>]*?)(\\s*\\/?)>`, 'i');
  const match = pattern.exec(xml);
  const updateTag = (qualifiedName: string, oldAttributes: string, selfClose: string) => {
    const attributes = new Map<string, string>();
    const attributePattern = /(?:^|\s)([\w:.-]+)\s*=\s*(["'])(.*?)\2/g;
    for (const item of oldAttributes.matchAll(attributePattern)) attributes.set(item[1]!, item[3]!);
    for (const [name, value] of Object.entries(wanted)) {
      if (value === undefined) attributes.delete(name);
      else attributes.set(name, value);
    }
    const serialized = [...attributes].map(([name, value]) => ` ${name}="${value.replaceAll('&', '&amp;').replaceAll('"', '&quot;')}"`).join('');
    return `<${qualifiedName}${serialized}${selfClose || '/'}>`;
  };
  if (match && match.index !== undefined) return xml.slice(0, match.index) + updateTag(match[1]!, match[2]!, match[3]!) + xml.slice(match.index + match[0].length);
  const tag = updateTag(elementName, '', '/');
  const insertion = insertBefore.exec(xml);
  if (!insertion || insertion.index === undefined) throw new Error(`No valid insertion point for ${elementName}.`);
  return xml.slice(0, insertion.index) + tag + xml.slice(insertion.index);
}

function setFitToPage(xml: string): string {
  const sheetPr = /<((?:[\w.-]+:)?sheetPr)\b([^>]*?)(\s*\/?)>/i.exec(xml);
  if (!sheetPr || sheetPr.index === undefined) {
    const anchor = /<(?:[\w.-]+:)?(?:dimension|sheetViews|sheetFormatPr|cols|sheetData)\b/i;
    const insertion = anchor.exec(xml);
    if (!insertion || insertion.index === undefined) throw new Error('Worksheet has no valid sheetPr insertion point.');
    return `${xml.slice(0, insertion.index)}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>${xml.slice(insertion.index)}`;
  }
  const name = sheetPr[1]!;
  if (sheetPr[3]) {
    const expanded = `<${name}${sheetPr[2]}><pageSetUpPr fitToPage="1"/></${name}>`;
    return xml.slice(0, sheetPr.index) + expanded + xml.slice(sheetPr.index + sheetPr[0].length);
  }
  const closePattern = new RegExp(`</${name.replace(':', '\\:')}\\s*>`, 'i');
  const close = closePattern.exec(xml.slice(sheetPr.index + sheetPr[0].length));
  if (!close || close.index === undefined) throw new Error('Unclosed worksheet sheetPr element.');
  const contentStart = sheetPr.index + sheetPr[0].length;
  const content = xml.slice(contentStart, contentStart + close.index);
  const setup = /<((?:[\w.-]+:)?pageSetUpPr)\b([^>]*?)(\s*\/?)>/i.exec(content);
  let updated: string;
  if (setup && setup.index !== undefined) {
    const tag = replaceElement(content, 'pageSetUpPr', { fitToPage: '1' }, /$^/);
    updated = tag;
  } else {
    updated = `${content}<pageSetUpPr fitToPage="1"/>`;
  }
  return xml.slice(0, contentStart) + updated + xml.slice(contentStart + close.index);
}

function sheetPath(workbookXml: string, relationshipsXml: string, sheetName: string): string {
  const workbook = parseXml(workbookXml, 'xl/workbook.xml');
  const relationships = parseXml(relationshipsXml, 'xl/_rels/workbook.xml.rels');
  const sheets = Array.from(workbook.getElementsByTagNameNS(MAIN_NS, 'sheet'));
  const sheet = sheets.find((item) => item.getAttribute('name') === sheetName);
  if (!sheet) throw new Error(`Worksheet '${sheetName}' does not exist.`);
  const relationshipId = sheet.getAttributeNS(DOC_REL_NS, 'id');
  const relationship = Array.from(relationships.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship'))
    .find((item) => item.getAttribute('Id') === relationshipId);
  const target = relationship?.getAttribute('Target');
  if (!target) throw new Error(`Worksheet '${sheetName}' has no workbook relationship.`);
  const normalized = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const segments: string[] = [];
  for (const part of normalized.replaceAll('\\', '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') segments.pop();
    else segments.push(part);
  }
  const path = segments.join('/');
  if (!path.startsWith('xl/worksheets/') || path.includes('../')) throw new Error(`Worksheet relationship escapes the workbook package: ${path}`);
  return path;
}

function patchWorksheet(source: string, change: PrintLayoutChange, path: string): string {
  parseXml(source, path);
  let xml = setFitToPage(source);
  const setupOrder = /<(?:(?:[\w.-]+):)?(?:headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/(?:[\w.-]+:)?worksheet\s*>/i;
  xml = replaceElement(xml, 'pageSetup', {
    orientation: change.orientation,
    fitToWidth: String(change.fitToWidth ?? 1),
    fitToHeight: String(change.fitToHeight ?? 0),
    scale: undefined,
  }, setupOrder);
  parseXml(xml, path);
  const orientation = getElementAttribute(xml, 'pageSetup', 'orientation');
  const width = getElementAttribute(xml, 'pageSetup', 'fitToWidth');
  const height = getElementAttribute(xml, 'pageSetup', 'fitToHeight');
  const fit = getElementAttribute(xml, 'pageSetUpPr', 'fitToPage');
  if (orientation !== (change.orientation ?? getElementAttribute(source, 'pageSetup', 'orientation'))
    || width !== String(change.fitToWidth ?? 1) || height !== String(change.fitToHeight ?? 0) || fit !== '1') {
    throw new Error(`Print setup did not round-trip in ${path}.`);
  }
  if (normalizePrintSettings(source) !== normalizePrintSettings(xml)) {
    throw new Error(`Worksheet content outside the requested print setup changed in ${path}.`);
  }
  return xml;
}

function normalizePrintSettings(xml: string): string {
  return xml
    .replace(/<((?:[\w.-]+:)?pageSetup)\b([^>]*?)(?:\s*\/?)>/gi, (_whole, name: string, rawAttributes: string) => {
      const attributes = rawAttributes.replace(/(?:^|\s)(?:orientation|fitToWidth|fitToHeight|scale)\s*=\s*(["']).*?\1/gi, '');
      return attributes.trim() ? `<${name}${attributes}/>` : '';
    })
    .replace(/<((?:[\w.-]+:)?pageSetUpPr)\b([^>]*?)(?:\s*\/?)>/gi, (_whole, name: string, rawAttributes: string) => {
      const attributes = rawAttributes.replace(/(?:^|\s)fitToPage\s*=\s*(["']).*?\1/gi, '');
      return attributes.trim() ? `<${name}${attributes}/>` : '';
    })
    .replace(/<((?:[\w.-]+:)?sheetPr)\b([^>]*?)\s*\/\s*>/gi, (_whole, name: string, attributes: string) =>
      attributes.trim() ? `<${name}${attributes}/>` : '')
    .replace(/<((?:[\w.-]+:)?sheetPr)\b([^>]*)>\s*<\/\1\s*>/gi, (_whole, name: string, attributes: string) =>
      attributes.trim() ? `<${name}${attributes}/>` : '');
}

export function applyPrintLayout(bytes: Uint8Array, changes: PrintLayoutChange[]) {
  const entries = unzipSync(bytes, { filter: (entry) => !entry.name.endsWith('/') });
  const workbookXml = entries['xl/workbook.xml'];
  const relationshipsXml = entries['xl/_rels/workbook.xml.rels'];
  if (!workbookXml || !relationshipsXml) throw new Error('XLSX workbook parts are missing.');
  const workbookText = strFromU8(workbookXml);
  const relationshipsText = strFromU8(relationshipsXml);
  const paths = new Set<string>();
  const changedSheets = changes.map((change) => {
    const path = sheetPath(workbookText, relationshipsText, change.sheet);
    if (paths.has(path)) throw new Error(`Worksheet '${change.sheet}' was selected more than once.`);
    paths.add(path);
    const original = entries[path];
    if (!original) throw new Error(`Worksheet package part is missing: ${path}.`);
    const output = strToU8(patchWorksheet(strFromU8(original), change, path));
    entries[path] = output;
    return { sheet: change.sheet, part: path, afterBytes: output.length };
  });
  const output = zipSync(entries, { level: 6 });
  const reopened = unzipSync(output, { filter: (entry) => !entry.name.endsWith('/') });
  if (Object.keys(entries).sort().join('\n') !== Object.keys(reopened).sort().join('\n')) throw new Error('XLSX package parts changed during print layout update.');
  for (const [path, data] of Object.entries(entries)) {
    if (!paths.has(path) && (!reopened[path] || data.length !== reopened[path]!.length || data.some((byte, index) => byte !== reopened[path]![index]))) {
      throw new Error(`Untouched XLSX part changed: ${path}.`);
    }
  }
  for (const change of changedSheets) {
    const part = reopened[change.part];
    if (!part) throw new Error(`Changed worksheet part is missing after ZIP rewrite: ${change.part}.`);
    parseXml(strFromU8(part), change.part);
  }
  return { bytes: output, changedSheets, sheetContentPreserved: true as const };
}

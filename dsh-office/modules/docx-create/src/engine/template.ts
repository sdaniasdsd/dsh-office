import type { WordprocessingMLPackage } from '@docx4j/core-ts';
import { XmlPart } from '@docx4j/core-ts/parts';
import { failure } from '../errors';
import { parseXml, serializeXml, W } from './xml';
const token = /\{\{([A-Za-z][A-Za-z0-9_]{0,63})\}\}/g;

/** Text substitution only, no expression evaluator. Changed parts are DOM-serialized, not object-model rebuilt. */
export async function fillTemplate(pkg: WordprocessingMLPackage, values: Record<string, string>): Promise<{ filledKeys: string[]; changedParts: string[] }> {
  const used = new Set<string>();
  const changedParts: string[] = [];
  for (const part of pkg.parts) {
    if (!(part instanceof XmlPart)) continue;
    const name = part.partName.name;
    if (!name.startsWith('/word/')) continue;
    const original = await part.getXml();
    const doc = parseXml(original);
    if (doc.getElementsByTagNameNS(W, 'documentProtection').length || doc.getElementsByTagNameNS(W, 'trackRevisions').length || doc.getElementsByTagNameNS(W, 'permStart').length)
      failure('UNSUPPORTED_TEMPLATE', 'Protected or revision-tracked templates require an explicit docx-edit strategy.');
    if (!/^\/word\/(?:document|header\d*|footer\d*)\.xml$/.test(name)) {
      if ((doc.documentElement?.textContent ?? '').includes('{{')) failure('UNSUPPORTED_TEMPLATE', 'Placeholders are only supported in the body, tables, headers and footers.');
      continue;
    }
    let changed = false;
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W, 'p'));
    for (const p of paragraphs) {
      const nodes = Array.from(p.getElementsByTagNameNS(W, 't'));
      const joined = nodes.map(n => n.textContent ?? '').join('');
      if (!joined.includes('{{') && !joined.includes('}}')) continue;
      const matches = [...joined.matchAll(token)];
      if (!matches.length || joined.replace(token, '').includes('{{') || joined.replace(token, '').includes('}}')) failure('UNSUPPORTED_TEMPLATE', 'Malformed placeholder; use {{name}} within one paragraph.');
      const starts: number[] = []; let offset = 0;
      for (const n of nodes) { starts.push(offset); offset += (n.textContent ?? '').length; }
      for (const match of matches.reverse()) {
        const key = match[1]!;
        if (!Object.hasOwn(values, key)) failure('TEMPLATE_VALUE_MISSING', `Missing template value: ${key}`);
        const replacement = values[key]!;
        // Multi-line insertion belongs to structured creation in v1, not an ambiguous run rewrite.
        if (/[\r\n\t]/.test(replacement)) failure('UNSUPPORTED_TEMPLATE', 'Template values must be single-line text; use structured creation for multi-line content.');
        const start = match.index!; const end = start + match[0].length;
        const touched = nodes.filter((n, i) => starts[i]! < end && starts[i]! + (n.textContent ?? '').length > start);
        if (touched.some(n => n.parentNode?.parentNode !== p)) failure('UNSUPPORTED_TEMPLATE', 'Placeholder is inside nested content or a field.');
        const children = Array.from(p.childNodes);
        const first = children.indexOf(touched[0]!.parentNode!);
        const last = children.indexOf(touched[touched.length - 1]!.parentNode!);
        for (const child of children.slice(first, last + 1)) {
          if (child.nodeType !== 1) continue;
          const element = child as typeof p;
          if (element.namespaceURI !== W || element.localName !== 'r' || Array.from(child.childNodes).some(n => n.nodeType === 1 && ((n as typeof p).namespaceURI !== W || !['t', 'rPr'].includes((n as typeof p).localName ?? ''))))
            failure('UNSUPPORTED_TEMPLATE', 'Placeholder crosses a field, line break or other non-text content.');
        }
        for (let i = nodes.length - 1; i >= 0; i--) {
          const node = nodes[i]!; const base = starts[i]!;
          const length = (node.textContent ?? '').length;
          if (base >= end || base + length <= start) continue;
          const text = node.textContent ?? '';
          node.textContent = text.slice(0, Math.max(0, start - base)) + (start >= base ? replacement : '') + text.slice(Math.max(0, end - base));
          node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
        }
        used.add(key); changed = true;
      }
    }
    // Catch placeholders in unsupported XML locations rather than silently ignoring them.
    const beforeTokens = [...(doc.documentElement?.textContent ?? '').matchAll(token)];
    if (!changed && beforeTokens.length) failure('UNSUPPORTED_TEMPLATE', 'Placeholder is outside a supported text paragraph.');
    if (changed) { part.setXml(serializeXml(doc)); changedParts.push(part.partName.storeName); }
  }
  for (const key of Object.keys(values)) if (!used.has(key)) failure('TEMPLATE_VALUE_UNUSED', `Template does not use value: ${key}`);
  if (!used.size) failure('UNSUPPORTED_TEMPLATE', 'No supported placeholders found.');
  return { filledKeys: [...used].sort(), changedParts };
}

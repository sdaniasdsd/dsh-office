import { unzipSync, strFromU8 } from 'fflate';
import { posix } from 'node:path';
import type { ModuleConfig } from '../contract';
import type { SafetyPolicy } from 'office-safety';
import { failure, DocxCreateError } from '../errors';
import { parseXml, W, W14, R } from './xml';
export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export interface PackageFacts { paragraphIds: string[]; paragraphs: number; tables: number; fields: boolean; entries: number }
export function checkPackage(bytes: Uint8Array, config: ModuleConfig, policy?: SafetyPolicy): PackageFacts {
  if (!bytes.length || bytes.length > Math.max(config.limits.maxInputBytes, config.limits.maxOutputBytes)) failure('LIMIT_EXCEEDED', 'Package byte limit exceeded.');
  let total = 0;
  const names = new Set<string>();
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(bytes, { filter: file => {
      if (names.has(file.name) || file.name.startsWith('/') || file.name.includes('\\') || file.name.split('/').includes('..'))
        failure('FORMAT_MISMATCH', 'Unsafe or duplicate ZIP part name.');
      names.add(file.name); total += file.originalSize;
      if (names.size > config.limits.maxEntries || total > config.limits.maxExpandedBytes) failure('LIMIT_EXCEEDED', 'Expanded package exceeds configured limits.');
      return true;
    } });
  } catch (error) {
    if (error instanceof DocxCreateError) throw error;
    return failure('FORMAT_MISMATCH', 'Not a supported, unencrypted DOCX ZIP package.');
  }
  if (!entries['word/document.xml'] || !entries['[Content_Types].xml'] || !entries['_rels/.rels']) failure('FORMAT_MISMATCH', 'Required DOCX parts are missing.');
  const docs = new Map<string, ReturnType<typeof parseXml>>();
  for (const [name, data] of Object.entries(entries)) {
    if (/vbaProject|activeX|embeddings\/|altChunk|_xmlsignatures\//i.test(name)) failure('SAFETY_POLICY_DENIED', 'Macros, embedded programs, signed packages and active content are not supported.');
    if (name.endsWith('.xml') || name.endsWith('.rels')) docs.set(name, parseXml(strFromU8(data)));
  }
  const types = docs.get('[Content_Types].xml')!;
  const overrides = Array.from(types.getElementsByTagNameNS(CT_NS, 'Override'));
  if (!overrides.some(e => e.getAttribute('PartName') === '/word/document.xml' && e.getAttribute('ContentType') === MAIN_TYPE))
    failure('FORMAT_MISMATCH', 'Only DOCX main document content type is supported.');
  if (strFromU8(entries['[Content_Types].xml']!).match(/macroEnabled|vbaProject|activeX|oleObject/i)) failure('SAFETY_POLICY_DENIED', 'Active content type is forbidden.');
  const relations = new Map<string, Map<string, string>>();
  for (const [name, doc] of docs) if (name.endsWith('.rels')) {
    const source = name === '_rels/.rels' ? '' : name.replace(/(^|\/)\_rels\//, '$1').replace(/\.rels$/, '');
    if (doc.documentElement?.namespaceURI !== REL_NS) failure('FORMAT_MISMATCH', 'Invalid relationships namespace.');
    const ids = new Map<string, string>();
    for (const rel of Array.from(doc.getElementsByTagNameNS(REL_NS, 'Relationship'))) {
      const id = rel.getAttribute('Id'); const target = rel.getAttribute('Target');
      if (!id || !target || ids.has(id)) failure('FORMAT_MISMATCH', 'Invalid relationship ID or target.');
      // Baseline is intentionally stronger than an allowExternalLinks=true policy.
      if (rel.getAttribute('TargetMode') === 'External' || /^(?:[a-z]+:|\/\/)/i.test(target)) failure('SAFETY_POLICY_DENIED', 'External relationships are not supported in this version.');
      let decoded: string;
      try { decoded = decodeURIComponent(target.split('#')[0]!); } catch { return failure('FORMAT_MISMATCH', 'Malformed relationship URI.'); }
      const resolved = decoded.startsWith('/') ? decoded.slice(1) : posix.normalize(posix.join(posix.dirname(source), decoded));
      if (!entries[resolved]) failure('VERIFICATION_FAILED', 'An internal relationship target is missing.');
      if (/aFChunk|oleObject|attachedTemplate|vbaProject|control$/i.test(rel.getAttribute('Type') ?? '')) failure('SAFETY_POLICY_DENIED', 'Unsupported active relationship.');
      ids.set(id, resolved);
      if (!source && /\/officeDocument$/.test(rel.getAttribute('Type') ?? '') && resolved !== 'word/document.xml') failure('FORMAT_MISMATCH', 'Unsupported main document location.');
    }
    relations.set(source, ids);
  }
  if (![...(relations.get('')?.values() ?? [])].includes('word/document.xml')) failure('FORMAT_MISMATCH', 'Main document relationship is missing.');
  let paragraphs = 0; let tables = 0; let fields = false;
  const paragraphIds: string[] = []; const allIds = new Set<string>();
  for (const [name, doc] of docs) if (!name.endsWith('.rels')) {
    const list = Array.from(doc.getElementsByTagNameNS(W, 'p')); paragraphs += list.length;
    tables += doc.getElementsByTagNameNS(W, 'tbl').length;
    if (name === 'word/document.xml' && (doc.documentElement?.namespaceURI !== W || doc.documentElement.localName !== 'document' || doc.getElementsByTagNameNS(W, 'body').length !== 1)) failure('FORMAT_MISMATCH', 'Invalid document body.');
    for (const para of list) {
      const id = para.getAttributeNS(W14, 'paraId');
      if (id) {
        if (!/^[0-7][0-9A-Fa-f]{7}$/.test(id) || id === '00000000' || allIds.has(id.toUpperCase())) failure('VERIFICATION_FAILED', 'Invalid or duplicate paragraph IDs.');
        allIds.add(id.toUpperCase()); paragraphIds.push(id.toUpperCase());
      }
    }
    for (const table of Array.from(doc.getElementsByTagNameNS(W, 'tbl'))) {
      if (table.getElementsByTagNameNS(W, 'tc').length > Math.min(config.limits.maxTableCells, policy?.maxTableCells ?? Infinity)) failure('LIMIT_EXCEEDED', 'Table cell limit exceeded.');
    }
    const instructions = Array.from(doc.getElementsByTagNameNS(W, 'instrText')).map(e => e.textContent).join(' ');
    const simpleFields = Array.from(doc.getElementsByTagNameNS(W, 'fldSimple'));
    const code = instructions + ' ' + simpleFields.map(e => e.getAttributeNS(W, 'instr')).join(' ');
    if (/\b(?:DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|LINK|DATABASE)\b/i.test(code)) failure('SAFETY_POLICY_DENIED', 'External or active field instructions are not supported.');
    fields ||= !!instructions.trim() || simpleFields.length > 0 || doc.getElementsByTagNameNS(W, 'fldChar').length > 0;
    for (const el of Array.from(doc.getElementsByTagName('*'))) {
      for (const attr of Array.from(el.attributes)) if (attr.namespaceURI === R && ['id', 'embed', 'link'].includes(attr.localName ?? '')) {
        if (!relations.get(name)?.has(attr.value)) failure('VERIFICATION_FAILED', 'A document element references a missing relationship.');
      }
    }
  }
  if (paragraphs + tables > config.limits.maxBlocks) failure('LIMIT_EXCEEDED', 'Document block limit exceeded.');
  const body = docs.get('word/document.xml')!.getElementsByTagNameNS(W, 'body')[0]!;
  const bodyBlocks = Array.from(body.childNodes).filter(n => n.nodeType === 1 && ['p', 'tbl'].includes((n as typeof body).localName ?? '')).length;
  if (policy?.maxBlocks !== undefined && bodyBlocks > policy.maxBlocks) failure('SAFETY_POLICY_DENIED', 'Policy block limit exceeded.');
  return { paragraphIds, paragraphs, tables, fields, entries: names.size };
}

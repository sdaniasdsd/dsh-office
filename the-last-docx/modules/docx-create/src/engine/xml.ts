import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { failure } from '../errors';
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`;
export function xml(text: string) { return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;'); }
export function parseXml(source: string) {
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) failure('SAFETY_POLICY_DENIED', 'DTD and entity declarations are forbidden.');
  try {
    return new DOMParser({ onError: () => { throw new Error('Malformed XML'); } }).parseFromString(source, 'application/xml');
  } catch { return failure('FORMAT_MISMATCH', 'A package XML part is malformed.'); }
}
export function serializeXml(document: ReturnType<typeof parseXml>) { return new XMLSerializer().serializeToString(document); }
/** Preserve line breaks and tabs as Word elements, not whitespace in w:t. */
export function textRun(text: string, properties = ''): string {
  const chunks = text.replace(/\r\n?/g, '\n').split(/([\n\t])/);
  return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}${chunks.map(chunk => chunk === '\n' ? '<w:br/>' : chunk === '\t' ? '<w:tab/>' : `<w:t xml:space="preserve">${xml(chunk)}</w:t>`).join('')}</w:r>`;
}

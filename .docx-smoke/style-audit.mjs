// Style-reference audit: which named styles a DOCX *references* versus which it
// actually *defines*. A pStyle pointing at an undefined styleId is a dangling
// reference: the document opens, but the named style carries no appearance.
import { readFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';

function read(file, member) {
  const bytes = readFileSync(file);
  let i = 0;
  while (i < bytes.length - 4) {
    if (bytes.readUInt32LE(i) !== 0x04034b50) { i += 1; continue; }
    const flags = bytes.readUInt16LE(i + 6);
    const nameLength = bytes.readUInt16LE(i + 26);
    const extraLength = bytes.readUInt16LE(i + 28);
    const compressedSize = bytes.readUInt32LE(i + 18);
    const name = bytes.subarray(i + 30, i + 30 + nameLength).toString('utf8');
    const dataStart = i + 30 + nameLength + extraLength;
    if (name === member && (flags & 0x08) === 0) {
      return inflateRawSync(bytes.subarray(dataStart, dataStart + compressedSize)).toString('utf8');
    }
    i = dataStart + compressedSize;
  }
  return undefined;
}

const [, , label, file] = process.argv;
const styles = read(file, 'word/styles.xml') ?? '';
const document = read(file, 'word/document.xml') ?? '';

const defined = [...styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"/g)].map((m) => m[1]);
const referenced = [...new Set([...document.matchAll(/<w:pStyle w:val="([^"]+)"/g)].map((m) => m[1]))];
const dangling = referenced.filter((id) => !defined.includes(id));

// How much appearance a defined style actually carries.
const described = new Map();
for (const match of styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
  const body = match[2];
  const has = (token) => (body.includes(token) ? token : null);
  const traits = ['<w:b/>', '<w:color ', '<w:sz ', '<w:spacing ', '<w:pBdr>', '<w:outlineLvl '].filter(has);
  described.set(match[1], traits.join(' '));
}

console.log(`### ${label}`);
console.log(`styles.xml: ${styles.length}B   document.xml: ${document.length}B`);
console.log(`defined  (${defined.length}): ${defined.join(', ') || '(none)'}`);
console.log(`referenced(${referenced.length}): ${referenced.join(', ') || '(none)'}`);
console.log(`dangling  (${dangling.length}): ${dangling.join(', ') || '(none)'}`);
for (const [id, traits] of described) console.log(`  ${id.padEnd(14)} carries: ${traits || '(nothing but the name)'}`);
console.log();

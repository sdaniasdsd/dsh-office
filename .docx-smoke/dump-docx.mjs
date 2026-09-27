// Dump a DOCX package: every part with its uncompressed size, plus the full
// text of selected members so two packages can be compared structurally.
import { readFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';

const [, , path, ...wanted] = process.argv;

function members(file) {
  const bytes = readFileSync(file);
  const out = new Map();
  let i = 0;
  while (i < bytes.length - 4) {
    if (bytes.readUInt32LE(i) !== 0x04034b50) { i += 1; continue; }
    const flags = bytes.readUInt16LE(i + 6);
    const nameLength = bytes.readUInt16LE(i + 26);
    const extraLength = bytes.readUInt16LE(i + 28);
    const compressedSize = bytes.readUInt32LE(i + 18);
    const uncompressedSize = bytes.readUInt32LE(i + 22);
    const name = bytes.subarray(i + 30, i + 30 + nameLength).toString('utf8');
    const dataStart = i + 30 + nameLength + extraLength;
    const raw = (flags & 0x08) === 0
      ? inflateRawSync(bytes.subarray(dataStart, dataStart + compressedSize))
      : Buffer.alloc(uncompressedSize);
    out.set(name, raw);
    i = dataStart + compressedSize;
  }
  return out;
}

const parts = members(path);
const fileSize = readFileSync(path).length;
console.log(`### ${path}`);
console.log(`container=${fileSize}B  parts=${parts.size}`);
for (const [name, data] of [...parts].sort((a, b) => a[0].localeCompare(b[0]))) {
  console.log(`  ${String(data.length).padStart(7)}B  ${name}`);
}
for (const name of wanted) {
  const data = parts.get(name);
  console.log(`\n--- ${name} ${data ? `(${data.length}B)` : '(absent)'} ---`);
  if (data) console.log(data.toString('utf8'));
}

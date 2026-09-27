// Compare the inner members of two DOCX packages, so a regenerated fixture can
// be told apart from a real content change. ZIP bytes differ every run because
// the generator writes fresh timestamps; the payload is what matters.
import { readFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';
import { createHash } from 'node:crypto';

function members(path) {
  const bytes = readFileSync(path);
  const out = new Map();
  let i = 0;
  while (i < bytes.length - 4) {
    if (bytes.readUInt32LE(i) !== 0x04034b50) { i += 1; continue; }
    const flags = bytes.readUInt16LE(i + 6);
    const nameLength = bytes.readUInt16LE(i + 26);
    const extraLength = bytes.readUInt16LE(i + 28);
    const compressedSize = bytes.readUInt32LE(i + 18);
    const name = bytes.subarray(i + 30, i + 30 + nameLength).toString('utf8');
    const dataStart = i + 30 + nameLength + extraLength;
    if ((flags & 0x08) === 0) {
      out.set(name, inflateRawSync(bytes.subarray(dataStart, dataStart + compressedSize)));
    }
    i = dataStart + compressedSize;
  }
  return out;
}

const [, , left, right] = process.argv;
const a = members(left);
const b = members(right);
console.log(`members: ${a.size} vs ${b.size}`);
const names = [...new Set([...a.keys(), ...b.keys()])].sort();
let differing = 0;
for (const name of names) {
  const ha = a.has(name) ? createHash('sha256').update(a.get(name)).digest('hex') : 'MISSING';
  const hb = b.has(name) ? createHash('sha256').update(b.get(name)).digest('hex') : 'MISSING';
  const same = ha === hb;
  if (!same) differing += 1;
  console.log(`${same ? 'same' : 'DIFF'}  ${name}  ${ha.slice(0, 12)}  ${hb.slice(0, 12)}`);
}
console.log(differing === 0 ? 'INNER CONTENT IDENTICAL' : `INNER CONTENT DIFFERS in ${differing} member(s)`);
process.exit(differing === 0 ? 0 : 1);

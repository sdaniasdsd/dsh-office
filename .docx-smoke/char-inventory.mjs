// List every non-ASCII character a document uses, with its codepoint and count.
//
// A "decoding problem" and a "wrong character" look the same in a screenshot but
// are different faults: the first is a pipeline losing a codepoint, the second is
// the plan choosing a different glyph. This tells them apart.
//
//   node char-inventory.mjs <file.json>   (JSON: [{label, file}])
import { readFileSync } from 'node:fs';

// A BOM is tolerated on purpose: this file is often written by PowerShell, whose
// -Encoding UTF8 emits one, and a BOM makes JSON.parse throw.
const items = JSON.parse(readFileSync(process.argv[2], 'utf8').replace(/^\uFEFF/u, ''));
const reports = [];

for (const { label, file } of items) {
  const text = readFileSync(file, 'utf8');
  // Only text that a reader sees: the contents of <w:t> elements.
  const visible = [...text.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/gu)].map((m) => m[1]).join('');
  const counts = new Map();
  for (const char of [...visible].filter((c) => c.codePointAt(0) > 127)) {
    counts.set(char, (counts.get(char) ?? 0) + 1);
  }
  reports.push({ label, total: visible.length, counts });
}

const all = new Set();
for (const r of reports) for (const char of r.counts.keys()) all.add(char);

const sorted = [...all].sort((a, b) => (a.codePointAt(0) ?? 0) - (b.codePointAt(0) ?? 0));
console.log(`characters seen in either document: ${sorted.length}\n`);
console.log('char  codepoint   ' + reports.map((r) => r.label.padStart(12)).join(''));
for (const char of sorted) {
  const codepoint = `U+${(char.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}`;
  const cells = reports.map((r) => String(r.counts.get(char) ?? 0).padStart(12)).join('');
  const only = reports.filter((r) => (r.counts.get(char) ?? 0) > 0).length === 1 ? '   <-- only in one' : '';
  console.log(`  ${char}   ${codepoint}   ${cells}${only}`);
}

console.log('\nsummary per document:');
for (const r of reports) {
  const fills = [...r.counts].filter(([c]) => [0x005f, 0xff3f, 0x2017, 0x2500].includes(c.codePointAt(0)));
  console.log(`  ${r.label}: ${r.total} visible chars, ${r.counts.size} distinct non-ASCII`);
  console.log(`     fill characters: ${fills.map(([c, n]) => `${c} U+${(c.codePointAt(0) ?? 0).toString(16).toUpperCase()} x${n}`).join(', ') || '(none)'}`);
}

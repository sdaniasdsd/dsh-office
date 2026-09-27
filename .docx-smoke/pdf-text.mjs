// Extract rendered text from a PDF with pdftotext and compare it against expected
// strings. All literals live in this UTF-8 file on purpose: passing CJK literals
// through a pwsh `-Command` string mangles them, which produced a whole batch of
// false MISS results in an earlier probe.
//
//   node pdf-text.mjs <pdftotext.exe> <pdf> [expected.json]
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { tmpdir } from 'node:os';

const PDFTOTEXT = process.argv[2];
const PDF = process.argv[3];
const EXPECTED_JSON = process.argv[4];

const squash = (value) => String(value).replace(/\s+/gu, '');

const txt = join(tmpdir(), `pdftext-${process.pid}.txt`);
execFileSync(PDFTOTEXT, [PDF, txt], { stdio: 'inherit' });
const raw = readFileSync(txt, 'utf8');
rmSync(txt, { force: true });

// Split on the form feed pdftotext writes at the end of every page.
const pages = raw.split('\f').filter((page, index, all) => !(index === all.length - 1 && page.trim() === ''));
const flat = squash(raw);

console.log(`pages with content: ${pages.length}`);
for (const [index, page] of pages.entries()) {
  const body = squash(page);
  console.log(`  page ${index + 1}: ${body.length} chars${body.length < 20 ? '   <-- suspiciously empty' : ''}`);
}
console.log(`\n--- full rendered text (whitespace squashed, ${flat.length} chars) ---`);
console.log(flat);
console.log('--- end ---\n');

if (EXPECTED_JSON) {
  const expected = JSON.parse(readFileSync(EXPECTED_JSON, 'utf8'));
  let missing = 0;
  for (const probe of expected.anywhere ?? []) {
    const ok = flat.includes(squash(probe));
    if (!ok) missing += 1;
    console.log(`${ok ? 'OK  ' : 'MISS'}  ${probe}`);
  }
  console.log(`\nmissing: ${missing}/${(expected.anywhere ?? []).length}`);
  process.exit(missing === 0 ? 0 : 1);
}

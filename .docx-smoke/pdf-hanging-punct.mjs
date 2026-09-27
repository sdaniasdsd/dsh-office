// Is the character crossing the right margin a punctuation mark? Hanging
// punctuation is normal Chinese typesetting; an overflowing ideograph is not.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, PDF] = process.argv.slice(2);
const xml = join(tmpdir(), `hang-${process.pid}.xml`);
execFileSync(PDFTOTEXT, ['-bbox-layout', PDF, xml], { stdio: 'inherit' });
const raw = readFileSync(xml, 'utf8');
rmSync(xml, { force: true });

const rows = [];
for (const page of raw.matchAll(/<page width="([\d.]+)"[\s\S]*?<\/page>/gu)) {
  const width = Number(page[1]);
  for (const line of page[0].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
    const xMax = Number(/xMax="([\d.]+)"/u.exec(line[1])?.[1]);
    const words = [...line[2].matchAll(/<word\b([^>]*)>([^<]*)<\/word>/gu)];
    if (!words.length || !Number.isFinite(xMax)) continue;
    const last = words[words.length - 1];
    const lastXMax = Number(/xMax="([\d.]+)"/u.exec(last[1])?.[1]);
    const lastXMin = Number(/xMin="([\d.]+)"/u.exec(last[1])?.[1]);
    rows.push({
      xMax, width,
      lastChar: last[2],
      lastWidthPt: Math.round((lastXMax - lastXMin) * 10) / 10,
      text: words.map((w) => w[2]).join(''),
    });
  }
}

const PUNCTUATION = /[。，、；：？！）】》”’·…—～%]$/u;
const mm = (r, v) => Math.round((v / r.width) * 210 * 10) / 10;
console.log('widest lines, with the character at the right edge:');
for (const r of [...rows].sort((a, b) => b.xMax - a.xMax).slice(0, 6)) {
  const edge = mm(r, r.xMax);
  const last = r.lastChar.slice(-1);
  console.log(`  ${String(edge).padStart(6)}mm  last="${last}" width=${r.lastWidthPt}pt  ${PUNCTUATION.test(last) ? 'punctuation (may hang)' : 'IDEOGRAPH'}   "${r.text.slice(-16)}"`);
}
const hanging = rows.filter((r) => mm(r, r.xMax) > 182.5);
console.log(`\nlines past 182.5 mm: ${hanging.length}`);
for (const r of hanging) console.log(`  ${mm(r, r.xMax)}mm  last char "${r.lastChar.slice(-1)}"  ${PUNCTUATION.test(r.lastChar.slice(-1)) ? 'OK - hanging punctuation' : 'NOT punctuation'}`);

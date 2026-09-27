// Dump a page of a PDF line by line, with position and glyph height, so the
// layout *shape* of two documents can be compared without seeing them.
//
//   node pdf-page-shape.mjs <pdftotext.exe> <pdf> <page> <label>
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, PDF, PAGE, LABEL] = process.argv.slice(2);
const xml = join(tmpdir(), `shape-${process.pid}.xml`);
execFileSync(PDFTOTEXT, ['-bbox-layout', PDF, xml], { stdio: 'inherit' });
const raw = readFileSync(xml, 'utf8');
rmSync(xml, { force: true });

const pages = [...raw.matchAll(/<page width="([\d.]+)" height="([\d.]+)"[\s\S]*?<\/page>/gu)];
const page = pages[Number(PAGE) - 1];
if (!page) { console.log(`${LABEL}: no page ${PAGE}`); process.exit(0); }
const width = Number(page[1]);
const toMm = (v) => Math.round((v / width) * 210 * 10) / 10;

console.log(`\n===== ${LABEL} — page ${PAGE} of ${pages.length} (A4) =====`);
const rows = [];
for (const line of page[0].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
  const xMin = Number(/xMin="([\d.]+)"/u.exec(line[1])?.[1]);
  const xMax = Number(/xMax="([\d.]+)"/u.exec(line[1])?.[1]);
  const yMin = Number(/yMin="([\d.]+)"/u.exec(line[1])?.[1]);
  const yMax = Number(/yMax="([\d.]+)"/u.exec(line[1])?.[1]);
  const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('');
  if (!text.trim()) continue;
  rows.push({ yMin, xMin, xMax, h: yMax - yMin, text: text.replace(/\s+/gu, '') });
}
rows.sort((a, b) => a.yMin - b.yMin);
for (const r of rows) {
  console.log(`  y=${String(Math.round(r.yMin)).padStart(4)}  左=${String(toMm(r.xMin)).padStart(6)}mm  宽=${String(Math.round(toMm(r.xMax) - toMm(r.xMin))).padStart(3)}mm  字高=${String(Math.round(r.h * 10) / 10).padStart(5)}pt  ${r.text.slice(0, 46)}`);
}
console.log(`  (${rows.length} lines)`);

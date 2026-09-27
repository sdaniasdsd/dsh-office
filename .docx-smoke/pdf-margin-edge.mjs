// Which line reaches furthest right, and does it cross the margin?
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, PDF, ...margins] = process.argv.slice(2);
const [leftMm, rightMm] = margins.map(Number);

const xml = join(tmpdir(), `edge-${process.pid}.xml`);
execFileSync(PDFTOTEXT, ['-bbox-layout', PDF, xml], { stdio: 'inherit' });
const raw = readFileSync(xml, 'utf8');
rmSync(xml, { force: true });

const rows = [];
for (const page of raw.matchAll(/<page width="([\d.]+)"[\s\S]*?<\/page>/gu)) {
  const width = Number(page[1]);
  for (const line of page[0].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
    const box = line[1];
    const xMax = Number(/xMax="([\d.]+)"/u.exec(box)?.[1]);
    const xMin = Number(/xMin="([\d.]+)"/u.exec(box)?.[1]);
    const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('').replace(/\s+/gu, '');
    if (!text || !Number.isFinite(xMax)) continue;
    rows.push({ xMin, xMax, text, width, mm: (v) => (v / width) * 210 });
  }
}
const limit = (210 - rightMm);
console.log(`right margin at ${limit} mm (page ${210} mm, right margin ${rightMm} mm)`);
console.log(`left margin at ${leftMm} mm\n`);
console.log('lines reaching furthest right:');
for (const r of [...rows].sort((a, b) => b.xMax - a.xMax).slice(0, 8)) {
  const mm = Math.round(r.mm(r.xMax) * 10) / 10;
  console.log(`  xMax=${String(Math.round(r.xMax * 10) / 10).padStart(6)}pt = ${String(mm).padStart(6)}mm  ${mm > limit + 0.5 ? 'BEYOND MARGIN' : 'ok            '}  "${r.text.slice(0, 40)}"`);
}
console.log('\nlines starting left of the left margin:');
for (const r of [...rows].sort((a, b) => a.xMin - b.xMin).slice(0, 5)) {
  const mm = Math.round(r.mm(r.xMin) * 10) / 10;
  console.log(`  xMin=${String(Math.round(r.xMin * 10) / 10).padStart(6)}pt = ${String(mm).padStart(6)}mm  ${mm < leftMm - 0.5 ? 'LEFT OF MARGIN' : 'ok            '}  "${r.text.slice(0, 40)}"`);
}

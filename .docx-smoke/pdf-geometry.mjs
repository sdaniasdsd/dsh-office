// Measure a rendered PDF's text block: left edge, right edge and line pitch, so
// a page-geometry claim can be checked against the page rather than the XML.
//
//   node pdf-geometry.mjs <pdftotext.exe> <a.pdf> <label> [<b.pdf> <label> ...]
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, ...rest] = process.argv.slice(2);
const inputs = [];
for (let i = 0; i + 1 < rest.length; i += 2) inputs.push({ pdf: rest[i], label: rest[i + 1] });

function measure(pdf) {
  const xml = join(tmpdir(), `geo-${process.pid}-${Math.random().toString(36).slice(2)}.xml`);
  execFileSync(PDFTOTEXT, ['-bbox-layout', pdf, xml], { stdio: 'inherit' });
  const raw = readFileSync(xml, 'utf8');
  rmSync(xml, { force: true });

  let pageWidth = 0;
  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  const pitches = [];
  for (const page of raw.matchAll(/<page width="([\d.]+)"[\s\S]*?<\/page>/gu)) {
    pageWidth = Number(page[1]);
    let previousY = null;
    for (const line of page[0].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
      const box = line[1];
      const xMin = Number(/xMin="([\d.]+)"/u.exec(box)?.[1]);
      const xMax = Number(/xMax="([\d.]+)"/u.exec(box)?.[1]);
      const yMin = Number(/yMin="([\d.]+)"/u.exec(box)?.[1]);
      const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('').replace(/\s+/gu, '');
      if (!text || !Number.isFinite(xMin)) continue;
      minX = Math.min(minX, xMin);
      maxX = Math.max(maxX, xMax);
      if (previousY !== null && yMin > previousY && yMin - previousY < 30) pitches.push(Math.round((yMin - previousY) * 10) / 10);
      previousY = yMin;
    }
  }
  const sorted = [...pitches].sort((a, b) => a - b);
  return { pageWidth, minX, maxX, medianPitch: sorted[Math.floor(sorted.length / 2)] };
}

const toMm = (points, pageWidth) => Math.round((points / pageWidth) * 210 * 10) / 10;

for (const { pdf, label } of inputs) {
  const r = measure(pdf);
  console.log(`  ${label.padEnd(18)} 左 ${String(r.minX).padStart(6)}pt (${toMm(r.minX, r.pageWidth)}mm)   距右 ${String(Math.round((r.pageWidth - r.maxX) * 10) / 10).padStart(6)}pt (${toMm(r.pageWidth - r.maxX, r.pageWidth)}mm)   行距中位 ${r.medianPitch}pt`);
}

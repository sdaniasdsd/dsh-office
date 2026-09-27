// Measure rendered line spacing from a PDF, without looking at it.
//
// `pdftotext -bbox-layout` reports a bounding box per line and per word, so the
// vertical pitch between consecutive lines is measurable directly. That turns
// "the line spacing looks wrong" into a number, and a number can be compared.
//
//   node pdf-line-pitch.mjs <pdftotext.exe> <a.pdf> <a.label> <b.pdf> <b.label>
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const PDFTOTEXT = process.argv[2];
const inputs = [];
for (let i = 3; i + 1 < process.argv.length; i += 2) inputs.push({ pdf: process.argv[i], label: process.argv[i + 1] });

const round = (value) => Math.round(value * 10) / 10;

function linesOf(pdf) {
  const xml = join(tmpdir(), `pitch-${process.pid}-${Math.random().toString(36).slice(2)}.xml`);
  execFileSync(PDFTOTEXT, ['-bbox-layout', pdf, xml], { stdio: 'inherit' });
  const raw = readFileSync(xml, 'utf8');
  rmSync(xml, { force: true });
  const pages = [];
  const pageRe = /<page\b[^>]*>([\s\S]*?)<\/page>/gu;
  for (const page of raw.matchAll(pageRe)) {
    const lines = [];
    for (const line of page[1].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
      const box = line[1];
      const yMin = Number(/yMin="([\d.]+)"/u.exec(box)?.[1]);
      const yMax = Number(/yMax="([\d.]+)"/u.exec(box)?.[1]);
      const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('');
      if (Number.isFinite(yMin) && Number.isFinite(yMax)) lines.push({ yMin, yMax, height: yMax - yMin, text });
    }
    pages.push(lines.sort((a, b) => a.yMin - b.yMin));
  }
  return pages;
}

function histogram(values) {
  const counts = new Map();
  for (const v of values) counts.set(v, (counts.get(v) ?? 0) + 1);
  return [...counts].sort((a, b) => b[1] - a[1]);
}

for (const { pdf, label } of inputs) {
  const pages = linesOf(pdf);
  const heights = [];
  const byPair = new Map();
  const gaps = [];
  for (const lines of pages) {
    for (let i = 0; i < lines.length; i += 1) {
      heights.push(round(lines[i].height));
      if (i === 0) continue;
      const pitch = round(lines[i].yMin - lines[i - 1].yMin);
      const prevH = round(lines[i - 1].height);
      const thisH = round(lines[i].height);
      if (pitch > 0 && pitch < 30 && Math.abs(prevH - thisH) < 1.5) {
        const key = `${prevH} pt -> ${thisH} pt`;
        if (!byPair.has(key)) byPair.set(key, []);
        byPair.get(key).push(pitch);
      }
      if (pitch >= 10 && pitch < 90) gaps.push(pitch);
    }
  }
  console.log(`\n================ ${label} ================`);
  console.log(`pages=${pages.length}`);
  console.log('\n-- line pitch between two lines of the SAME glyph height (comparable text) --');
  const rows = [...byPair].map(([key, values]) => {
    const sorted = [...values].sort((a, b) => a - b);
    return { key, n: sorted.length, median: sorted[Math.floor(sorted.length / 2)] };
  }).filter((r) => r.n >= 3).sort((a, b) => b.n - a.n);
  for (const r of rows.slice(0, 8)) console.log(`  ${String(r.n).padStart(4)}x  ${r.key.padEnd(18)} median pitch = ${r.median} pt`);

  console.log('\n-- spacing between consecutive paragraphs (line pitch above the normal pitch) --');
  const body = gaps.filter((g) => g > 10 && g < 30);
  const para = gaps.filter((g) => g >= 30 && g < 90);
  const hist = (list) => {
    const counts = new Map();
    for (const v of list) counts.set(v, (counts.get(v) ?? 0) + 1);
    return [...counts].sort((a, b) => b[1] - a[1]).slice(0, 6);
  };
  console.log('  paragraph-to-paragraph extra gaps: ' + (para.length ? hist(para).map(([v, c]) => `${c}x ${v}pt`).join('  ') : '(none above 30 pt)'));
  console.log(`  body-pitch candidates: ${body.length ? hist(body).map(([v, c]) => `${c}x ${v}pt`).join('  ') : '(none)'}`);
}

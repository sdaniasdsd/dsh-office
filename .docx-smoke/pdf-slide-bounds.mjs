// Per-page text bounding box, to check no slide's text runs off the page.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, PDF] = process.argv.slice(2);
const xml = join(tmpdir(), `bounds-${process.pid}.xml`);
execFileSync(PDFTOTEXT, ['-bbox-layout', PDF, xml], { stdio: 'inherit' });
const raw = readFileSync(xml, 'utf8');
rmSync(xml, { force: true });

const pages = [...raw.matchAll(/<page width="([\d.]+)" height="([\d.]+)"[\s\S]*?<\/page>/gu)];
console.log(`pages: ${pages.length}`);
let worst = null;
for (const [index, page] of pages.entries()) {
  const width = Number(page[1]);
  const height = Number(page[2]);
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
  for (const word of page[0].matchAll(/<word\b([^>]*)>([^<]*)<\/word>/gu)) {
    if (!word[2].trim()) continue;
    const xMin = Number(/xMin="([\d.]+)"/u.exec(word[1])?.[1]);
    const xMax = Number(/xMax="([\d.]+)"/u.exec(word[1])?.[1]);
    const yMin = Number(/yMin="([\d.]+)"/u.exec(word[1])?.[1]);
    const yMax = Number(/yMax="([\d.]+)"/u.exec(word[1])?.[1]);
    if (![xMin, xMax, yMin, yMax].every(Number.isFinite)) continue;
    minX = Math.min(minX, xMin); maxX = Math.max(maxX, xMax);
    minY = Math.min(minY, yMin); maxY = Math.max(maxY, yMax);
  }
  const overflow = maxX > width || maxY > height || minX < 0 || minY < 0;
  const row = { page: index + 1, leftPt: Math.round(minX), rightGapPt: Math.round(width - maxX), topPt: Math.round(minY), bottomGapPt: Math.round(height - maxY), overflow };
  if (overflow) worst = row;
  console.log(`  p${String(row.page).padStart(2)}  左 ${String(row.leftPt).padStart(4)}pt  距右 ${String(row.rightGapPt).padStart(4)}pt  上 ${String(row.topPt).padStart(4)}pt  距下 ${String(row.bottomGapPt).padStart(4)}pt ${overflow ? '  ← 越界' : ''}`);
}
console.log(worst ? `\n有页面越界：${JSON.stringify(worst)}` : '\n所有页面文字都在页面内');

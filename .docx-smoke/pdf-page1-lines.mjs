// Page 1, line by line: where each line starts and how tall it is, so a page-1
// layout can be compared against the standard without looking at images.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, ...rest] = process.argv.slice(2);
const inputs = [];
for (let i = 0; i + 1 < rest.length; i += 2) inputs.push({ pdf: rest[i], label: rest[i + 1] });
const PAGE = Number(process.env.PAGE ?? 1);

for (const { pdf, label } of inputs) {
  const xml = join(tmpdir(), `p1-${process.pid}-${Math.random().toString(36).slice(2)}.xml`);
  execFileSync(PDFTOTEXT, ['-bbox-layout', pdf, xml], { stdio: 'inherit' });
  const raw = readFileSync(xml, 'utf8');
  rmSync(xml, { force: true });
  const pages = [...raw.matchAll(/<page width="([\d.]+)" height="([\d.]+)"[\s\S]*?<\/page>/gu)];
  const page = pages[PAGE - 1];
  if (!page) { console.log(`\n===== ${label}: no page ${PAGE} =====`); continue; }
  const width = Number(page[1]);
  const height = Number(page[2]);
  const mm = (value) => Math.round((value / width) * 210 * 10) / 10;
  const mmY = (value) => Math.round((value / height) * 297 * 10) / 10;
  console.log(`\n===== ${label} — page ${PAGE} of ${pages.length} =====`);
  let index = 0;
  for (const line of page[0].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
    const box = line[1];
    const xMin = Number(/xMin="([\d.]+)"/u.exec(box)?.[1]);
    const xMax = Number(/xMax="([\d.]+)"/u.exec(box)?.[1]);
    const yMin = Number(/yMin="([\d.]+)"/u.exec(box)?.[1]);
    const yMax = Number(/yMax="([\d.]+)"/u.exec(box)?.[1]);
    const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('').replace(/\s+/gu, '');
    if (!text || !Number.isFinite(xMin)) continue;
    index += 1;
    const limit = Number(process.env.LIMIT ?? 40);
    console.log(`  ${String(index).padStart(2)}  y ${String(mmY(yMin)).padStart(6)}mm  x ${String(mm(xMin)).padStart(6)}→${String(mm(xMax)).padStart(6)}mm  字高 ${String(Math.round((yMax - yMin) * 10) / 10).padStart(5)}pt  [${[...text].length}字] ${text.slice(0, limit)}`);
  }
}

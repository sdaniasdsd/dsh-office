// Which glyph sizes does each document actually render, and which text carries
// the odd ones? Aggregated so a size inconsistency is a number, not an impression.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const [PDFTOTEXT, ...rest] = process.argv.slice(2);
const inputs = [];
for (let i = 0; i + 1 < rest.length; i += 2) inputs.push({ pdf: rest[i], label: rest[i + 1] });

for (const { pdf, label } of inputs) {
  const xml = join(tmpdir(), `sizes-${process.pid}-${Math.random().toString(36).slice(2)}.xml`);
  execFileSync(PDFTOTEXT, ['-bbox-layout', pdf, xml], { stdio: 'inherit' });
  const raw = readFileSync(xml, 'utf8');
  rmSync(xml, { force: true });

  const bySize = new Map();
  const samples = new Map();
  for (const page of raw.matchAll(/<page\b[^>]*>([\s\S]*?)<\/page>/gu)) {
    for (const line of page[1].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
      const yMax = Number(/yMax="([\d.]+)"/u.exec(line[1])?.[1]);
      const yMin = Number(/yMin="([\d.]+)"/u.exec(line[1])?.[1]);
      const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('').replace(/\s+/gu, '');
      if (!text || !Number.isFinite(yMax)) continue;
      const size = Math.round((yMax - yMin) * 10) / 10;
      bySize.set(size, (bySize.get(size) ?? 0) + 1);
      if (!samples.has(size)) samples.set(size, []);
      if (samples.get(size).length < 3) samples.get(size).push(text.slice(0, 26));
    }
  }
  console.log(`\n===== ${label} =====`);
  for (const [size, count] of [...bySize].sort((a, b) => b[0] - a[0])) {
    console.log(`  字高 ${String(size).padStart(5)} pt  ×${String(count).padStart(4)}   e.g. ${samples.get(size).join(' | ')}`);
  }
}

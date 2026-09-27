// Measure the row pitch of the form tables in one PDF, by locating the lines that
// carry known field labels and comparing their y positions.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const PDFTOTEXT = process.argv[2];
const PDF = process.argv[3];
const LABELS = ['单位名称', '住', '法定代表人', '统一社会信用代码', '联系电话', '姓', '性', '身份证件号码', '户籍地址', '现居住地址', '甲方（盖章）', '乙方（签名）'];

const xml = join(tmpdir(), `rows-${process.pid}.xml`);
execFileSync(PDFTOTEXT, ['-bbox-layout', PDF, xml], { stdio: 'inherit' });
const raw = readFileSync(xml, 'utf8');
rmSync(xml, { force: true });

const lines = [];
for (const page of raw.matchAll(/<page\b[^>]*>([\s\S]*?)<\/page>/gu)) {
  for (const line of page[1].matchAll(/<line\b([^>]*)>([\s\S]*?)<\/line>/gu)) {
    const yMin = Number(/yMin="([\d.]+)"/u.exec(line[1])?.[1]);
    const yMax = Number(/yMax="([\d.]+)"/u.exec(line[1])?.[1]);
    const text = [...line[2].matchAll(/<word\b[^>]*>([^<]*)<\/word>/gu)].map((w) => w[1]).join('');
    lines.push({ yMin, height: yMax - yMin, text });
  }
  lines.push({ yMin: NaN, height: NaN, text: '--- page break ---' });
}

console.log('lines carrying a form label, with the pitch to the previous labelled line:');
let previous = null;
for (const line of lines) {
  const flat = line.text.replace(/\s+/gu, '');
  if (!LABELS.some((l) => flat.startsWith(l.replace(/\s+/gu, '')))) continue;
  const pitch = previous === null || !Number.isFinite(line.yMin) ? '-' : Math.round((line.yMin - previous) * 10) / 10;
  console.log(`  y=${String(Math.round(line.yMin)).padStart(4)}  h=${String(Math.round(line.height * 10) / 10).padStart(5)}  pitch=${String(pitch).padStart(6)}  "${line.text.slice(0, 34)}"`);
  previous = line.yMin;
}

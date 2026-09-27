// Why does the same paragraph break one character differently in two documents?
// Dumps one paragraph's properties and text from each, codepoint by codepoint, so
// an invisible difference is a fact rather than a guess.
import { readFileSync } from 'node:fs';

const START = process.env.STARTS_WITH ?? '根据';
for (const path of process.argv.slice(2)) {
  const xml = readFileSync(path, 'utf8');
  const paragraphs = [...xml.matchAll(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/gu)].map((match) => match[0]);
  const paragraph = paragraphs.find((candidate) => [...candidate.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gu)].map((m) => m[1]).join('').startsWith(START));
  console.log(`\n===== ${path} =====`);
  if (!paragraph) { console.log(`  no paragraph starting with ${START}`); continue; }
  const ppr = /<w:pPr>[\s\S]*?<\/w:pPr>/u.exec(paragraph)?.[0] ?? '(no pPr)';
  const text = [...paragraph.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gu)].map((m) => m[1]).join('');
  const rprs = [...paragraph.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/gu)].map((m) => m[0]);
  console.log(`  pPr  : ${ppr}`);
  console.log(`  rPr  : ${rprs.join(' ')}`);
  console.log(`  chars: ${[...text].length}`);
  console.log(`  text : ${text}`);
  console.log(`  odd  : ${[...text].map((c) => (c.codePointAt(0) < 128 || c.codePointAt(0) > 0x2fff ? `${c}=U+${c.codePointAt(0).toString(16).toUpperCase()}` : '')).filter(Boolean).join(' ') || '(none)'}`);
}

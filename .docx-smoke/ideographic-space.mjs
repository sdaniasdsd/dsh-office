// Does a specific line keep its ideographic space? Counted on the XML text runs
// of two documents, so the answer does not depend on a PDF extractor's whitespace
// handling.
import { readFileSync } from 'node:fs';

const IDEOGRAPHIC_SPACE = '\u3000';
for (const path of process.argv.slice(2)) {
  const xml = readFileSync(path, 'utf8');
  const texts = [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gu)].map((match) => match[1]);
  const spaces = texts.filter((text) => text.includes(IDEOGRAPHIC_SPACE));
  const heading = texts.find((text) => text.startsWith('第一条'));
  console.log(`\n${path}`);
  console.log(`  text runs                : ${texts.length}`);
  console.log(`  runs with U+3000         : ${spaces.length}`);
  console.log(`  total U+3000             : ${(xml.match(/\u3000/gu) ?? []).length}`);
  console.log(`  heading run              : ${JSON.stringify(heading)}  codepoints=${[...(heading ?? '')].map((c) => c.codePointAt(0)).join(',')}`);
  const sample = spaces.slice(0, 3).map((text) => JSON.stringify(text));
  console.log(`  samples                  : ${sample.join(' | ')}`);
}

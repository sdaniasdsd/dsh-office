// Compare the *effective* formatting of two DOCX documents.
//
// A styles.xml diff is not evidence of what a reader sees: a paragraph can carry
// a built-in style and still be formatted by direct run properties, and an indent
// may be written as `w:firstLine` (twips) or `w:firstLineChars` (hundredths of a
// character). This reads the body and reports, per paragraph, the style it
// references, its own pPr, and the direct rPr its runs carry.
//
//   node docx-effective-format.mjs <a.document.xml> <a.styles.xml> <a.label> <b.document.xml> <b.styles.xml> <b.label>
import { readFileSync } from 'node:fs';

const [docA, styA, labelA, docB, styB, labelB] = process.argv.slice(2);
const read = (path) => readFileSync(path, 'utf8');

function stylesOf(xml) {
  const map = new Map();
  for (const m of xml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/gu)) {
    map.set(m[1], m[2]);
  }
  return map;
}
const attr = (fragment, pattern) => pattern.exec(fragment)?.[1];
const onOff = (fragment, name) => (new RegExp(`<w:${name}\\b[^>]*w:val="0"`, 'u').test(fragment) ? 'off' : new RegExp(`<w:${name}\\b`, 'u').test(fragment) ? 'on' : '-');

function paragraphsOf(xml) {
  const out = [];
  for (const m of xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/gu)) {
    const body = m[1];
    const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/u.exec(body)?.[1] ?? '';
    const text = [...body.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/gu)].map((t) => t[1]).join('');
    const firstRun = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/u.exec(body)?.[1] ?? '';
    const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/u.exec(firstRun)?.[1] ?? '';
    out.push({
      text,
      styleId: attr(pPr, /<w:pStyle w:val="([^"]+)"/u) ?? '(none)',
      jc: attr(pPr, /<w:jc w:val="([^"]+)"/u) ?? '-',
      firstLine: attr(pPr, /w:firstLine="(\d+)"/u) ?? '-',
      firstLineChars: attr(pPr, /w:firstLineChars="(\d+)"/u) ?? '-',
      spacing: attr(pPr, /<w:spacing\b([^>]*)\/>/u)?.trim() ?? '-',
      runEastAsia: attr(rPr, /w:eastAsia="([^"]+)"/u) ?? '-',
      runAscii: attr(rPr, /w:ascii="([^"]+)"/u) ?? '-',
      runSz: attr(rPr, /<w:sz w:val="(\d+)"/u) ?? '-',
      runBold: onOff(rPr, 'b'),
      runColor: attr(rPr, /<w:color w:val="([^"]+)"/u) ?? '-',
      hasDirectRPr: rPr.length > 0,
    });
  }
  return out;
}

function profile(label, docPath, styPath) {
  const xml = read(docPath);
  const styles = stylesOf(read(styPath));
  const paragraphs = paragraphsOf(xml);
  const body = paragraphs.filter((p) => p.text.trim().length > 0);

  console.log(`\n================ ${label} ================`);
  console.log(`paragraphs with text: ${body.length}`);
  const direct = body.filter((p) => p.hasDirectRPr).length;
  console.log(`paragraphs carrying direct run formatting: ${direct}/${body.length}`);

  const byStyle = new Map();
  for (const p of body) byStyle.set(p.styleId, (byStyle.get(p.styleId) ?? 0) + 1);
  console.log('styles referenced: ' + [...byStyle].map(([k, v]) => `${k}=${v}`).join('  '));

  // What the body actually looks like, as (style, direct size, direct bold, face, colour, alignment).
  const combos = new Map();
  for (const p of body) {
    const key = `${p.styleId} | sz=${p.runSz} b=${p.runBold} ea=${p.runEastAsia} color=${p.runColor} jc=${p.jc} firstLine=${p.firstLine}/${p.firstLineChars}`;
    combos.set(key, (combos.get(key) ?? 0) + 1);
  }
  console.log('\n-- effective paragraph formats (style + direct run props), most common first --');
  for (const [key, count] of [...combos].sort((a, b) => b[1] - a[1]).slice(0, 14)) console.log(`  ${String(count).padStart(3)}x  ${key}`);

  const interesting = body.filter((p) => p.text.startsWith('劳动合同书') || /^第[一二三四五六七八九十]+条/u.test(p.text.trim())).slice(0, 8);
  console.log('\n-- title and clause headings, verbatim --');
  for (const p of interesting) {
    const style = styles.get(p.styleId) ?? '';
    const styleSz = attr(style, /<w:sz w:val="(\d+)"/u) ?? '-';
    const styleEa = attr(style, /w:eastAsia="([^"]+)"/u) ?? '-';
    const styleJc = attr(style, /<w:jc w:val="([^"]+)"/u) ?? '-';
    console.log(`  "${p.text.slice(0, 22)}"`);
    console.log(`      pStyle=${p.styleId} (style sz=${styleSz} ea=${styleEa} jc=${styleJc})  jc=${p.jc} ind=${p.firstLine}/${p.firstLineChars} spacing="${p.spacing}"`);
    console.log(`      direct rPr: ea=${p.runEastAsia} ascii=${p.runAscii} sz=${p.runSz} bold=${p.runBold} color=${p.runColor}`);
  }
}

profile(labelA, docA, styA);
profile(labelB, docB, styB);

// Dump the table styles a mature OOXML library ships by default, so a table
// style can be modelled on real practice instead of invented.
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { writeFileSync } from 'node:fs';

const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4', defaultTheme: '2013' });
const bytes = await pkg.save();
writeFileSync(process.argv[2], bytes);

const styles = await pkg.getMainDocumentPart().styleDefinitionsPart?.getXml();
if (!styles) { console.log('no styles part'); process.exit(0); }

const tableStyles = [...styles.matchAll(/<w:style\b[^>]*w:type="table"[\s\S]*?<\/w:style>/g)].map((m) => m[0]);
console.log(`styles part: ${styles.length}B, table styles: ${tableStyles.length}`);
for (const style of tableStyles) {
  const id = /w:styleId="([^"]+)"/.exec(style)?.[1];
  console.log(`\n===== ${id} (${style.length}B) =====`);
  console.log(style.slice(0, 1800));
}

const grid = tableStyles.find((s) => /w:styleId="TableGrid"/.test(s))
  ?? tableStyles.find((s) => /w:styleId="TableNormal"/.test(s))
  ?? tableStyles[0];
console.log(`\n===== 参考样式全文 =====\n${grid ?? '(none)'}`);
const fonts = /<w:rFonts[^>]*\/>/.exec(styles)?.[0];
console.log(`\n===== docDefaults 字体 =====\n${fonts ?? '(none)'}`);
console.log(`\n===== docDefaults 段落 =====\n${/<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/.exec(styles)?.[0] ?? '(none)'}`);

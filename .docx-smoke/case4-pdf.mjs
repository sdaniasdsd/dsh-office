// Case 4's end-to-end rendering evidence.
//
// Tracked changes are rendered as displayed, so the pages legitimately contain
// BOTH the struck-through old text and the new text. What this checks is that
// no page went missing, none came back blank, and every approved value reached
// the rendered page.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const PDFTOTEXT = process.argv[2];
const OUT_PDF = process.argv[3];
const SRC_PDF = process.argv[4];

const squash = (value) => String(value).replace(/\s+/gu, '');
const pagesOf = (pdf) => {
  const txt = join(tmpdir(), `c4-${process.pid}-${Math.random().toString(36).slice(2)}.txt`);
  execFileSync(PDFTOTEXT, [pdf, txt], { stdio: 'inherit' });
  const raw = readFileSync(txt, 'utf8');
  rmSync(txt, { force: true });
  return raw.split('\f').filter((page, index, all) => !(index === all.length - 1 && page.trim() === ''));
};

const out = pagesOf(OUT_PDF);
const src = pagesOf(SRC_PDF);
console.log(`pages: source=${src.length} amended=${out.length}`);
const blanks = out.map((page, i) => [i + 1, squash(page).length]).filter(([, n]) => n < 40);
console.log(`blank or near-empty pages: ${blanks.length === 0 ? 'none' : JSON.stringify(blanks)}`);
console.log('per-page characters: ' + out.map((page, i) => `${i + 1}:${squash(page).length}`).join(' '));

const flat = squash(out.join('\n'));
const MUST = [
  ['2.1 start date', '2026年11月1日'],
  ['2.1 end date', '2029年10月31日'],
  ['2.2 three months', '试用期为三个月'],
  ['2.2 period end', '2027年1月31日'],
  ['3.1 role', '岗位为数据运营专员'],
  ['3.1 department', '所属部门为数据治理组'],
  ['4.1 location', '上海市浦东新区龙井路88号海悦中心5层'],
  ['6.1 base salary', '20800元'],
  ['6.1 written form', '贰万零捌佰元整'],
  ['6.2 probation salary', '16640元'],
  ['9.2 new sentence', '因故障排查形成的临时数据导出文件'],
  ['new comment', '请确认通讯地址、个人邮箱或联系电话变更时的书面通知及收悉记录均可追溯'],
  ['old comment kept', '请确认附件A与正文的工资项目仅作一致性复述'],
  ['footnote kept', '本段为虚构测试条款'],
  ['decoy transit allowance', '800元'],
  ['annex B emergency point', '上海市浦东新区云桥路168号澄海中心1层北门'],
];
let missing = 0;
console.log('');
for (const [label, needle] of MUST) {
  const ok = flat.includes(squash(needle));
  if (!ok) missing += 1;
  console.log(`${ok ? 'OK  ' : 'MISS'}  ${label.padEnd(26)} ${needle.slice(0, 40)}`);
}
console.log(`\nmissing: ${missing}/${MUST.length}`);
process.exit(missing === 0 && blanks.length === 0 && out.length === src.length ? 0 : 1);

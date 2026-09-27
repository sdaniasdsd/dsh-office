// Print context windows around anchors in a rendered PDF, so a "missing" probe
// can be told apart from a probe that simply assumed the wrong adjacency.
import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const PDFTOTEXT = process.argv[2];
const PDF = process.argv[3];

const txt = join(tmpdir(), `c4ctx-${process.pid}.txt`);
execFileSync(PDFTOTEXT, [PDF, txt], { stdio: 'inherit' });
const pages = readFileSync(txt, 'utf8').split('\f');
rmSync(txt, { force: true });

const ANCHORS = [
  '试用期为',
  '乙方岗位为',
  '所属部门为',
  '月基本工资为人民币',
  '试用期税前月工资',
  '日常工作地点为',
  '三个月培养回顾',
  '交通津贴为税前每月',
  '年度目标奖金参考额',
  '因故障排查形成的临时数据导出文件',
  '双方应保证本合同首页所列',
  '虚构测试条款',
  '工资项目仅作一致性复述',
  '通讯地址、个人邮箱或联系电话变更时的书面通知',
];

for (const anchor of ANCHORS) {
  let found = false;
  pages.forEach((page, index) => {
    const flat = page.replace(/\s+/gu, '');
    const at = flat.indexOf(anchor.replace(/\s+/gu, ''));
    if (at < 0) return;
    found = true;
    const from = Math.max(0, at - 30);
    console.log(`p${index + 1}  …${flat.slice(from, at + anchor.length + 90)}…`);
  });
  if (!found) console.log(`--   ${anchor}   NOT FOUND`);
  console.log('');
}

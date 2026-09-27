// Generate the labour contract to the approved sample's specification.
//
// Every format value here is read off the approved PDF/docx sample:
//   page      A4, 2.5 cm top and bottom, 2.8 cm outer, 3.0 cm gutter
//   title     黑体 二号 (22 pt) centred
//   clause    小四 (12 pt) 黑体 bold, 10 pt above / 4 pt below
//   body      宋体 小四, 1.5 leading, two-character first-line indent, 2 pt after
//   footer    centred 9 pt: 第 X 页 共 Y 页
//   tables    the signature block carries no edges
//
// Party-specific data stays blank: a standard form has fields, and inventing a
// company, a person or a salary would be fabricating facts.
import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { ArtifactRef } from 'office-core';
import { createDocxCreateModule } from './src/index';
import type { DocumentSpec } from './src/domain/docx-create';

const OUT = process.argv[2] ?? 'D:/contract';
// The output name is an argument, so a run never fights a file the reader has
// open: a locked target surfaces only as a bare store failure, which reads like
// a code fault and is not one.
const NAME = process.argv[3] ?? 'labour-contract-standard-form';
const B = '＿＿＿＿＿＿';
const D = '＿＿＿＿';
// A4 (210 mm) less the 30 mm gutter and the 28 mm outer margin leaves 152 mm.
const FORM: [number, number] = [55, 97];
const SIGN: [number, number] = [76, 76];

const blocks: DocumentSpec['blocks'] = [
  { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '劳动合同书' }] },

  { kind: 'paragraph', id: 'h-party-a', style: 'Heading2', runs: [{ text: '甲方（用人单位）' }] },
  { kind: 'table', id: 'party-a', columnWidthsMm: FORM, rows: [
    ['单位名称', B],
    ['住　　所', B],
    ['法定代表人（主要负责人）', B],
    ['统一社会信用代码', B],
    ['联系电话', B],
  ] },

  { kind: 'paragraph', id: 'h-party-b', style: 'Heading2', runs: [{ text: '乙方（劳动者）' }] },
  { kind: 'table', id: 'party-b', columnWidthsMm: FORM, rows: [
    ['姓　　名', B],
    ['性　　别', B],
    ['身份证件号码', B],
    ['户籍地址', B],
    ['现居住地址', B],
    ['联系电话', B],
  ] },

  { kind: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: '第一条　合同期限' }] },
  { kind: 'paragraph', id: 'p1-1', runs: [{ text: `1.1　本合同为固定期限劳动合同，自${D}年${D}月${D}日起至${D}年${D}月${D}日止。` }] },
  { kind: 'paragraph', id: 'p1-2', runs: [{ text: '1.2　本合同期限届满前，双方可以协商续订劳动合同。' }] },

  { kind: 'paragraph', id: 'h2', style: 'Heading1', runs: [{ text: '第二条　试用期' }] },
  { kind: 'paragraph', id: 'p2-1', runs: [{ text: `2.1　试用期为${B}个月，自${D}年${D}月${D}日起至${D}年${D}月${D}日止。试用期包含在本合同期限内。` }] },
  { kind: 'paragraph', id: 'p2-2', runs: [{ text: '2.2　乙方在试用期的工资不得低于本单位相同岗位最低档工资或者本合同约定工资的百分之八十，并不得低于甲方所在地最低工资标准。' }] },

  { kind: 'paragraph', id: 'h3', style: 'Heading1', runs: [{ text: '第三条　工作内容与工作地点' }] },
  { kind: 'paragraph', id: 'p3-1', runs: [{ text: `3.1　乙方的工作岗位为${B}，工作地点为${B}。` }] },
  { kind: 'paragraph', id: 'p3-2', runs: [{ text: '3.2　甲方可以根据生产经营需要及乙方的能力、工作表现，经与乙方协商一致后调整工作岗位和工作地点。' }] },

  { kind: 'paragraph', id: 'h4', style: 'Heading1', runs: [{ text: '第四条　工作时间与休息休假' }] },
  { kind: 'paragraph', id: 'p4-1', runs: [{ text: `4.1　甲方对乙方实行${B}工时制度（标准工时制／综合计算工时制／不定时工作制）。` }] },
  { kind: 'paragraph', id: 'p4-2', runs: [{ text: '4.2　乙方依法享有法定休假日、年休假及其他休假，甲方依法安排乙方休息休假。' }] },

  { kind: 'paragraph', id: 'h5', style: 'Heading1', runs: [{ text: '第五条　劳动报酬' }] },
  { kind: 'paragraph', id: 'p5-1', runs: [{ text: `5.1　乙方试用期月工资为人民币${B}元，试用期满后月工资为人民币${B}元。` }] },
  { kind: 'paragraph', id: 'p5-2', runs: [{ text: `5.2　甲方于每月${B}日前以货币形式支付乙方上月工资，并向乙方提供其本人的工资清单。` }] },
  { kind: 'paragraph', id: 'p5-3', runs: [{ text: '5.3　甲方依法代扣代缴个人所得税及乙方个人应当承担的社会保险、住房公积金部分。' }] },

  { kind: 'paragraph', id: 'h6', style: 'Heading1', runs: [{ text: '第六条　社会保险与福利' }] },
  { kind: 'paragraph', id: 'p6-1', runs: [{ text: '6.1　甲方依法为乙方办理社会保险登记手续，并按时足额缴纳社会保险费。' }] },
  { kind: 'paragraph', id: 'p6-2', runs: [{ text: '6.2　乙方依法享受甲方依法制定的福利待遇。' }] },

  { kind: 'paragraph', id: 'h7', style: 'Heading1', runs: [{ text: '第七条　劳动保护、劳动条件和职业危害防护' }] },
  { kind: 'paragraph', id: 'p7-1', runs: [{ text: '7.1　甲方为乙方提供符合国家规定的劳动安全卫生条件和必要的劳动防护用品。' }] },
  { kind: 'paragraph', id: 'p7-2', runs: [{ text: '7.2　乙方从事接触职业病危害作业的，甲方应当依法组织上岗前、在岗期间和离岗时的职业健康检查。' }] },

  { kind: 'paragraph', id: 'h8', style: 'Heading1', runs: [{ text: '第八条　劳动纪律与保密' }] },
  { kind: 'paragraph', id: 'p8-1', runs: [{ text: '8.1　乙方应当遵守甲方依法制定的规章制度和劳动纪律。' }] },
  { kind: 'paragraph', id: 'p8-2', runs: [{ text: '8.2　乙方应当对甲方的商业秘密以及与知识产权相关的保密事项履行保密义务。' }] },

  { kind: 'paragraph', id: 'h9', style: 'Heading1', runs: [{ text: '第九条　劳动合同的变更、解除与终止' }] },
  { kind: 'paragraph', id: 'p9-1', runs: [{ text: '9.1　双方协商一致，可以变更本合同约定的内容。变更劳动合同应当采用书面形式。' }] },
  { kind: 'paragraph', id: 'p9-2', runs: [{ text: '9.2　本合同的解除与终止，依照《中华人民共和国劳动合同法》及有关法律法规的规定执行。' }] },

  { kind: 'paragraph', id: 'h10', style: 'Heading1', runs: [{ text: '第十条　争议解决' }] },
  { kind: 'paragraph', id: 'p10-1', runs: [{ text: '10.1　因履行本合同发生争议，双方可以协商解决；协商不成的，可以依法申请调解、仲裁或者提起诉讼。' }] },

  { kind: 'paragraph', id: 'h11', style: 'Heading1', runs: [{ text: '第十一条　其他' }] },
  { kind: 'paragraph', id: 'p11-1', runs: [{ text: '11.1　本合同的附件与本合同具有同等法律效力。' }] },
  { kind: 'paragraph', id: 'p11-2', runs: [{ text: '11.2　本合同一式两份，甲乙双方各执一份，自双方签字（盖章）之日起生效。' }] },
  { kind: 'paragraph', id: 'p11-3', runs: [{ text: '11.3　本合同未尽事宜，依照国家及地方有关法律法规执行。' }] },

  { kind: 'paragraph', id: 'h-sign', style: 'Heading1', runs: [{ text: '签署' }] },
  // A layout table, not a data table: no edges.
  { kind: 'table', id: 'sign', borders: 'none', columnWidthsMm: SIGN, rows: [
    ['甲方（盖章）：', '乙方（签名）：'],
    ['法定代表人或委托代理人：', '身份证件号码：'],
    [`签署日期：${D}年${D}月${D}日`, `签署日期：${D}年${D}月${D}日`],
  ] },
];

await mkdir(OUT, { recursive: true });

const store = {
  async read(ref: ArtifactRef) { throw new Error(`read not needed: ${ref.id}`); },
  async write({ bytes, requestId }: { bytes: Uint8Array; requestId: string }): Promise<ArtifactRef> {
    const name = `${requestId}.docx`;
    const path = join(OUT, name);
    await writeFile(path, bytes);
    return { id: name, uri: path, label: name, sizeBytes: bytes.length,
      sha256: createHash('sha256').update(bytes).digest('hex'),
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
  },
};

const module = createDocxCreateModule({ artifactStore: store });
const out = await module.handlers.execute({
  requestId: NAME,
  operation: 'execute',
  plan: { kind: 'create', document: {
    preset: 'chinese-contract',
    scenario: 'formal-record',
    page: { size: 'A4', marginsMm: { top: 25, right: 28, bottom: 25, left: 30 } },
    pageNumberStyle: 'pageOfTotal',
    blocks,
  } },
});
const design = out.result.design!;
console.log(`register    : ${design.register}  [${design.label}]`);
console.log(`decoration  : ${design.decoration}`);
console.log(`decided by  : ${design.source}`);
console.log(`verification: ${out.verification?.ok ? 'ok' : 'FAILED'} ${out.verification?.checks?.filter((c) => c.status === 'pass').length}/${out.verification?.checks?.length}`);
console.log(`file        : ${out.result.artifactRef.uri}`);
console.log(`sha256      : ${out.result.artifactRef.sha256}`);
console.log(`size        : ${out.result.artifactRef.sizeBytes} B`);
await module.dispose();

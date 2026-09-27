import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { createDocxCreateModule, PRESETS } from '../src/index';
import type { Preset, DocumentSpec } from '../src/domain/docx-create';
import type { ArtifactRef } from 'office-core';

const root = fileURLToPath(new URL('./_generated/', import.meta.url));
await mkdir(root, { recursive: true });
// Each build gets a new directory: rerunning never overwrites an earlier sample or template.
const directory = join(root, randomUUID()); await mkdir(directory);
const module = createDocxCreateModule({ artifactStore: {
  async read(ref) { return new Uint8Array(await readFile(ref.uri)); },
  async write({ requestId, bytes }) {
    const uri = resolve(directory, `${requestId}.docx`);
    if (!uri.startsWith(resolve(directory) + sep)) throw new Error('Invalid sample name');
    await writeFile(uri, bytes, { flag: 'wx' });
    return { id: randomUUID(), uri, label: `${requestId}.docx` };
  },
} });
const artifacts: ArtifactRef[] = [];
try {
  for (const preset of Object.keys(PRESETS) as Preset[]) {
    const document: DocumentSpec = {
      preset, header: '{{organization}} · 文档示例', footer: '项目资料', blocks: [
        { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '{{title}}' }] },
        { kind: 'paragraph', id: 'subtitle', style: 'Subtitle', runs: [{ text: '创建、解析、编辑共用一份文档' }] },
        { kind: 'paragraph', id: 'summary-heading', style: 'Heading1', runs: [{ text: '一、概要' }] },
        { kind: 'paragraph', id: 'summary', runs: [{ text: '{{summary}}' }] },
        { kind: 'paragraph', id: 'scope-heading', style: 'Heading1', runs: [{ text: '二、交付范围' }] },
        { kind: 'table', id: 'scope', columnWidthsMm: [38, 90, 32], rows: [
          ['能力', '说明', '状态'], ['创建', '标题、段落、表格、页眉页脚', '已实现'],
          ['模板', '正文、表格、页眉页脚文字替换', '已实现'], ['排版', '由独立渲染模块完成逐页检查', '待验证'],
        ] },
        { kind: 'paragraph', id: 'next-heading', style: 'Heading1', runs: [{ text: '三、后续工作' }] },
        { kind: 'paragraph', id: 'next', runs: [{ text: '生成完成后，解析模块负责建立节点索引；渲染模块负责生成页面图片。发现问题后，编辑模块根据节点定位修改。' }] },
        ...(preset === 'technical' ? [{ kind: 'paragraph' as const, id: 'code', style: 'Code' as const, runs: [{ text: 'create(document)\nparse(artifactRef)\nrender(artifactRef)' }] }] : []),
        ...(preset === 'chinese-long' ? [
          { kind: 'pageBreak' as const, id: 'appendix-break' },
          { kind: 'paragraph' as const, id: 'appendix-title', style: 'Heading1' as const, runs: [{ text: '附录：长文档注意事项' }] },
          { kind: 'paragraph' as const, id: 'appendix-body', runs: [{ text: '字体和排版环境会影响换行。目录只生成字段，不虚构最终页码。正式交付前应在指定的排版环境中更新字段并核验每一页。' }] },
        ] : []),
      ],
    };
    const template = await module.handlers.execute({ operation: 'execute', requestId: `${preset}-template`, plan: { kind: 'create', document } });
    artifacts.push(template.result.artifactRef);
    const filled = await module.handlers.execute({ operation: 'execute', requestId: `${preset}-sample`, plan: {
      kind: 'fillTemplate', templateRef: template.result.artifactRef,
      values: { organization: '示例团队', title: `${PRESETS[preset].label} · 示例`, summary: '这是一份通过 docx-create 生成的原生 Word 文档，可直接进入现有解析、编辑和渲染流程。' },
    } });
    artifacts.push(filled.result.artifactRef);
  }
  console.log(JSON.stringify({ directory, visualReview: 'pending', artifacts }, null, 2));
} finally { await module.dispose(); }

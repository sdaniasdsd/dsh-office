// Scratch harness: render one brief through each decoration register so the
// register choice can be judged side by side instead of argued about in prose.
// Not part of the module; regenerate with `node scripts/assemble.mjs` if it is
// ever removed.
import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { ArtifactRef } from 'office-core';
import { createDocxCreateModule } from './src/index';
import type { DocumentSpec } from './src/domain/docx-create';

const OUT = process.argv[2] ?? 'D:/register-compare';

const BLOCKS: DocumentSpec['blocks'] = [
  { kind: 'paragraph', id: 'title', style: 'Title', runs: [{ text: '第三季度客户交付简报' }] },
  { kind: 'paragraph', id: 'date', style: 'Subtitle', runs: [{ text: '2026年9月30日' }] },
  { kind: 'paragraph', id: 'h-summary', style: 'Heading1', runs: [{ text: '执行摘要' }] },
  { kind: 'paragraph', id: 'p-summary', runs: [{ text: '本季度完成 18 项交付任务，按期完成 16 项；尚有 2 项延期，其中 1 项等待客户确认，1 项进入修复。' }] },
  { kind: 'paragraph', id: 'h-table', style: 'Heading1', runs: [{ text: '汇总' }] },
  {
    kind: 'table', id: 'totals', header: true, columnWidthsMm: [38, 22, 100],
    rows: [
      ['指标', '数量', '说明'],
      ['交付任务', '18', '本季度完成'],
      ['按期完成', '16', '已完成交付任务'],
      ['延期事项', '2', '1 项等待客户确认，1 项进入修复'],
    ],
  },
  { kind: 'paragraph', id: 'h-next', style: 'Heading1', runs: [{ text: '后续跟进' }] },
  // A real list, so the markers come from the numbering part rather than being
  // typed into the text.
  { kind: 'list', id: 'next-actions', items: ['取得客户确认。', '完成 1 项修复中的延期事项。'] },
];

/** One case per scenario, so the judgement is exercised rather than the override. */
const CASES = [
  { name: '1-internal-review', scenario: 'internal-review' as const, preset: 'report' as const },
  { name: '2-academic-report', scenario: 'academic-report' as const, preset: 'chinese-long' as const },
  { name: '3-formal-record', scenario: 'formal-record' as const, preset: 'report' as const },
  { name: '4-technical-spec', scenario: 'technical-spec' as const, preset: 'technical' as const },
  { name: '5-no-scenario', preset: 'report' as const },
];

await mkdir(OUT, { recursive: true });

const store = {
  async read(ref: ArtifactRef) { throw new Error(`read not needed: ${ref.id}`); },
  async write({ bytes, requestId }: { bytes: Uint8Array; requestId: string }): Promise<ArtifactRef> {
    const name = `${requestId}.docx`;
    const path = join(OUT, name);
    await writeFile(path, bytes);
    return {
      id: name, uri: path, label: name, sizeBytes: bytes.length,
      sha256: createHash('sha256').update(bytes).digest('hex'),
      mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    };
  },
};

const module = createDocxCreateModule({ artifactStore: store });
for (const item of CASES) {
  const out = await module.handlers.execute({
    requestId: item.name,
    operation: 'execute',
    plan: { kind: 'create', document: { preset: item.preset, ...(item.scenario ? { scenario: item.scenario } : {}), blocks: BLOCKS } },
  });
  const design = out.result.design!;
  console.log(`\n=== ${item.name}  (scenario=${item.scenario ?? 'none'}, preset=${item.preset})`);
  console.log(`  register   : ${design.register}  [${design.label}]`);
  console.log(`  decoration : ${design.decoration}`);
  console.log(`  decided by : ${design.source}`);
  console.log(`  file       : ${out.result.artifactRef.uri}  ${out.result.artifactRef.sizeBytes}B`);
  console.log(`  sha256     : ${out.result.artifactRef.sha256}`);
}
await module.dispose();

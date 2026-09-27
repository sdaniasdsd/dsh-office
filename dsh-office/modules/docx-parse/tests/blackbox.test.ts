/**
 * 黑盒测试 —— 只通过公开接口观察行为，不碰任何内部实现。
 *
 * 与已有的三层测试的区别：
 *   contract / edge-cases / regression 会直接 import mapper、sourcemap 等内部模块，
 *   属于「白盒」。本文件刻意只从 `src/index` 导入，扮演一个【外部消费者】：
 *   拿真实 docx 文件 → execute → 记账 → 判档 → 出简报，全程不看内部一眼。
 *
 * 这样测出的结论才等于「接进 Profile 之后实际会发生什么」。
 *
 * 报告会写成 `fixtures/_generated/blackbox-report.md`：
 *   终端在中文环境下会乱码，把报告落盘再用编辑器看，内容才是可读的。
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { afterAll, describe, expect, it } from 'vitest';

import {
  FORMAT_REVISION_INTENTION_THRESHOLD,
  computeNodeDigest,
  createDocxParseModule,
  createRevisionLedger,
  decideGovernance,
  diffDualIR,
  governanceDirective,
  renderAgentBrief,
} from '../src/index';
import type { DocxDualIR, RevisionIntent, RevisionLedger } from '../src/index';
import { FIXTURES_DIR, artifactRef, fixture } from './support';

/** 报告中累积的段落；在 afterAll 一次性落盘。 */
const report: string[] = [];
const say = (line = ''): void => {
  report.push(line);
};

const module = createDocxParseModule();

/** 通过公开 execute 接口把一份 docx 解析成双 IR。 */
async function executeIR(name: string): Promise<DocxDualIR> {
  const output = await module.handlers.execute({
    artifactRef: artifactRef(fixture(name)),
    operation: 'execute',
    requestId: `blackbox-${name}`,
  });
  return output.result.ir.content;
}

/** 造一条意图。 */
function intent(overrides: Partial<RevisionIntent> = {}): RevisionIntent {
  return {
    id: 'intent-1',
    requestId: 'req-1',
    source: 'user',
    stage: 'revision',
    kind: 'content',
    text: '（未填写）',
    targetIds: [],
    ...overrides,
  };
}

/** 把一份双 IR 的可点名目标压成一行摘要，便于肉眼核对。 */
function outlineOf(content: DocxDualIR): string[] {
  const lines: string[] = [];
  for (const block of content.semantic.blocks) {
    if (block.kind === 'heading') lines.push(`H${block.level} ${block.text}`);
    else if (block.kind === 'paragraph') lines.push(`  p  ${block.text}`);
    else lines.push(`  tbl ${block.rows.length}x${block.gridColumns}`);
  }
  return lines;
}

/** 把差异列表压成可读行。 */
function renderDiff(changes: ReturnType<typeof diffDualIR>): string[] {
  if (changes.length === 0) return ['（无变化）'];
  return changes.map((change) => {
    const fields = change.fields.length > 0 ? ` [${change.fields.join(',')}]` : '';
    if (change.change === 'added') return `+ added${fields}  ${change.after}`;
    if (change.change === 'removed') return `- removed${fields}  ${change.before}`;
    return `~ modified${fields}  ${change.before} → ${change.after}`;
  });
}

afterAll(() => {
  mkdirSync(FIXTURES_DIR, { recursive: true });
  writeFileSync(join(FIXTURES_DIR, 'blackbox-report.md'), report.join('\n'), 'utf8');
});

describe('blackbox: 公开接口能否正常解析真实文档', () => {
  it('把两份真实 docx 解析成双 IR，且视图数量与方案符合承诺', async () => {
    const v1 = await executeIR('rev-v1.docx');
    const v2 = await executeIR('rev-v2.docx');

    say('# 黑盒报告：docx-parse 双 IR + 修订账本 + 门槛判定');
    say();
    say('> 全程只调用 `src/index` 的公开导出；样本是真实 .docx 文件。');
    say();

    say('## 1. 解析结果');
    say();
    for (const [name, content] of [
      ['rev-v1.docx', v1],
      ['rev-v2.docx', v2],
    ] as const) {
      say(`### ${name}`);
      say(`- viewCount = ${content.viewCount}（承诺为 2）`);
      say(`- scheme = ${content.scheme}`);
      say(`- 主部件 = ${content.physical.mainPart}`);
      say(`- 语义块 ${content.semantic.blocks.length} 个，对应表 ${content.sourceMap.entries.length} 条`);
      say('- 正文：');
      for (const line of outlineOf(content)) say(`  - ${line}`);
      say();
    }

    expect(v1.viewCount).toBe(2);
    expect(v2.viewCount).toBe(2);
    expect(v1.scheme).toBe('composite-anchor-v1');
    // v2 比 v1 多一段（Hiring resumed），因此块数正好 +1。
    expect(v2.semantic.blocks.length).toBe(v1.semantic.blocks.length + 1);
  });

  it('同一份文件两次解析得到完全一致的 IR（确定性）', async () => {
    const first = await executeIR('rev-v1.docx');
    const second = await executeIR('rev-v1.docx');
    expect(JSON.stringify(second)).toBe(JSON.stringify(first));
  });
});

describe('blackbox: 真实新旧版本差异', () => {
  it('准确认出「改了文字 / 新增 / 只换样式 / 只改一格」，且不误报未动的内容', async () => {
    const v1 = await executeIR('rev-v1.docx');
    const v2 = await executeIR('rev-v2.docx');
    const changes = diffDualIR(v1, v2);

    say('## 2. 修订差异（v1 → v2）');
    say();
    say('上游对 v2 做了这些改动：Revenue 改数字、中途加一段、Costs 只换样式、表格改一格。');
    say();
    say('```');
    for (const line of renderDiff(changes)) say(line);
    say('```');
    say();

    // 恰好 4 处改动：文字 / 新增 / 样式 / 单元格文字。
    expect(changes).toHaveLength(4);
    expect(changes.filter((entry) => entry.change === 'added')).toHaveLength(1);
    expect(changes.filter((entry) => entry.change === 'modified')).toHaveLength(3);

    const textEdit = changes.find((entry) => entry.before === 'Revenue grew by 10 percent.');
    expect(textEdit?.fields).toEqual(['text']);
    expect(textEdit?.after).toBe('Revenue grew by 12 percent.');

    // 一个字没改、只换样式：这是通用文本 diff 会完全漏掉的那类改动。
    const styleEdit = changes.find((entry) => entry.before === 'Costs were flat.');
    expect(styleEdit?.fields).toEqual(['style']);

    // 只改了表格里的某一格，落点是那个单元格段落而不是整张表。
    const cellEdit = changes.find((entry) => entry.after === '135');
    expect(cellEdit?.kind).toBe('paragraph');
    expect(cellEdit?.pointers[0]).toContain('/w:tbl[1]/w:tr[1]/w:tc[2]/w:p[1]');

    // 未动的段落与未动的单元格都不得出现在报告里。
    const touched = changes.map((entry) => entry.after ?? entry.before).join(' ');
    expect(touched).not.toContain('Quarterly Report');
    expect(touched).not.toContain('Outlook is stable');
    expect(touched).not.toContain('120');
  });

  it('对照：有无 paraId 决定「插入一段」是否波及后面所有段落', async () => {
    const driftV1 = await executeIR('rev-drift-v1.docx');
    const driftV2 = await executeIR('rev-drift-v2.docx');
    const driftChanges = diffDualIR(driftV1, driftV2);

    const stableV1 = await executeIR('rev-v1.docx');
    const stableV2 = await executeIR('rev-v2.docx');
    const stableChanges = diffDualIR(stableV1, stableV2);

    say('## 3. 位置漂移的对照');
    say();
    say('两边都一样是「中途插入一段」，唯一的差别是段落有没有 w14:paraId。');
    say();
    say('### 没有 paraId —— 身份退化为位置，后续段落整体错位');
    say();
    say('```');
    for (const line of renderDiff(driftChanges)) say(line);
    say('```');
    say();
    say('注意 `Gamma → Beta` 这一行：这是错位造成的假象，Gamma 一个字都没改过。');
    say();
    say('### 有 paraId —— 只报出真实改动');
    say();
    say('```');
    for (const line of renderDiff(stableChanges)) say(line);
    say('```');
    say();

    // 无 paraId：插入一段导致 2 处内容被误判为「被改写」（另有 1 处真实新增）。
    expect(driftChanges).toHaveLength(3);
    expect(driftChanges.filter((entry) => entry.change === 'modified')).toHaveLength(2);
    expect(
      driftChanges.some((entry) => entry.before === 'Gamma' && entry.after === 'Beta'),
    ).toBe(true);

    // 有 paraId：4 处真实改动，且没有任何「删除」误报。
    expect(stableChanges).toHaveLength(4);
    expect(stableChanges.filter((entry) => entry.change === 'added')).toHaveLength(1);
    expect(stableChanges.every((entry) => entry.change !== 'removed')).toBe(true);
  });

  it('差异按文档顺序排列，第一行是最先被改动的地方', async () => {
    const stableChanges = diffDualIR(await executeIR('rev-v1.docx'), await executeIR('rev-v2.docx'));
    // v2 里的改动依次发生在：Revenue(文字) → Hiring(新增) → Costs(样式) → 表格单元格。
    const summaries = stableChanges.map((entry) => entry.after ?? entry.before ?? '');
    expect(summaries[0]).toBe('Revenue grew by 12 percent.');
    expect(summaries[summaries.length - 1]).toBe('135');
  });
});

describe('blackbox: 外部意图能否收得到并更新', () => {
  it('收一条意图、落账、算出差异、更新基线', async () => {
    const ledger: RevisionLedger = createRevisionLedger();
    const v1 = await executeIR('rev-v1.docx');
    const v2 = await executeIR('rev-v2.docx');

    // 第一次登记：把 v1 设为基线（用户还没提修改）。
    const baselineEntry = await ledger.record({
      intent: intent({ id: 'i1', text: '（初始基线，无用户指令）' }),
      after: v1,
    });
    // 第二次登记：用户提了修改，文档真的变了。
    const revisionEntry = await ledger.record({
      intent: intent({
        id: 'i2',
        text: '把营收数字更新成 12%，另外成本那段换个样式',
        kind: 'content',
      }),
      after: v2,
    });

    say('## 4. 修订账本（外部意图入口）');
    say();
    say(`- 基线登记：revision=${baselineEntry.revision}，产生新版本=${baselineEntry.producedVersion}，差异 ${baselineEntry.changes.length} 处`);
    say(`- 用户修改：revision=${revisionEntry.revision}，产生新版本=${revisionEntry.producedVersion}，差异 ${revisionEntry.changes.length} 处`);
    say(`- 用户原话被原样保留：「${revisionEntry.intent.text}」`);
    say(`- 前后指纹：${revisionEntry.beforeFingerprint?.slice(0, 12)}… → ${revisionEntry.afterFingerprint?.slice(0, 12)}…`);
    say();

    expect(revisionEntry.revision).toBe(2);
    expect(revisionEntry.producedVersion).toBe(true);
    expect(revisionEntry.changes).toHaveLength(4);
    // 原话不得被改写。
    expect(revisionEntry.intent.text).toContain('12%');
    expect(revisionEntry.beforeFingerprint).toBe(baselineEntry.afterFingerprint);
  });

  it('收得住「还没动手就先提要求」这种输入', async () => {
    const ledger = createRevisionLedger();
    const entry = await ledger.record({
      intent: intent({
        id: 'r1',
        stage: 'requirement',
        kind: 'format',
        text: '全篇用宋体小四，一级标题黑体三号',
      }),
    });

    say('### 只表达要求、还没产生新版本');
    say();
    say(`- 意图已入账：「${entry.intent.text}」`);
    say(`- producedVersion = ${entry.producedVersion}（文档此刻确实还没变）`);
    say(`- changes = ${entry.changes.length} 处（不伪造差异）`);
    say();

    expect(entry.producedVersion).toBe(false);
    expect(entry.changes).toEqual([]);
    expect((await ledger.intents())).toHaveLength(1);
  });

  it('拒绝畸形意图，且不污染只追加的账本', async () => {
    const ledger = createRevisionLedger();
    await expect(
      ledger.record({ intent: { ...intent(), stage: 'whenever' } as never }),
    ).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    expect(await ledger.history()).toHaveLength(0);
  });
});

describe('blackbox: 门槛判定（硬编码条件）', () => {
  it('四个典型场景各给出可解释的档位', async () => {
    const formatRequirement = intent({ stage: 'requirement', kind: 'format', text: '按公司模板' });
    const twoFormatEdits = Array.from({ length: FORMAT_REVISION_INTENTION_THRESHOLD }, (_, i) =>
      intent({ id: `f${i}`, kind: 'format', text: `字号再大点(${i})` }),
    );

    const cases = [
      { label: '普通办公、无任何严格信号', intents: [], context: {} },
      { label: '用户先给了格式要求', intents: [formatRequirement], context: {} },
      { label: '合同/公文类场景', intents: [], context: { strictTextSubmission: true } },
      { label: '格式修改意愿明显（连续两次）', intents: twoFormatEdits, context: {} },
      { label: '单次生成、没提格式', intents: [], context: { singleShot: true } },
      {
        label: '单次生成 + 用户反复改格式',
        intents: twoFormatEdits,
        context: { singleShot: true },
      },
      {
        label: '单次生成 + 但先给了格式要求',
        intents: [formatRequirement],
        context: { singleShot: true },
      },
    ];

    say('## 5. 门槛判定');
    say();
    say('| 场景 | 档位 | 理由 | 命中信号 |');
    say('| --- | --- | --- | --- |');

    const decisions = cases.map((entry) => ({
      ...entry,
      decision: decideGovernance(entry.intents, entry.context),
    }));
    for (const entry of decisions) {
      const signals = entry.decision.matchedSignals.join(', ') || '—';
      say(`| ${entry.label} | ${entry.decision.mode} | ${entry.decision.reason} | ${signals} |`);
    }
    say();

    const byLabel = new Map(decisions.map((entry) => [entry.label, entry.decision]));
    expect(byLabel.get('普通办公、无任何严格信号')?.mode).toBe('lite');
    expect(byLabel.get('用户先给了格式要求')?.mode).toBe('strict');
    expect(byLabel.get('合同/公文类场景')?.mode).toBe('strict');
    expect(byLabel.get('格式修改意愿明显（连续两次）')?.mode).toBe('strict');
    expect(byLabel.get('单次生成、没提格式')?.mode).toBe('lite');
    expect(byLabel.get('单次生成 + 用户反复改格式')?.mode).toBe('exempt');
    // 用户主动把格式说清楚 = 他自己在意，豁免让位。
    expect(byLabel.get('单次生成 + 但先给了格式要求')?.mode).toBe('strict');
  });

  it('每档纪律文案互不相同且可执行', () => {
    expect(new Set([governanceDirective('strict'), governanceDirective('lite'), governanceDirective('exempt')]).size).toBe(3);
  });
});

describe('blackbox: 交给 Agent 的判断依据', () => {
  it('渲染出一份 Agent 可直接消费的简报', async () => {
    const ledger = createRevisionLedger();
    const v1 = await executeIR('rev-v1.docx');
    const v2 = await executeIR('rev-v2.docx');

    await ledger.record({ intent: intent({ id: 'i1', text: '（基线）' }), after: v1 });
    await ledger.record({
      intent: intent({ id: 'i2', text: '把营收改成 12%，成本换样式，再加一段招聘进展' }),
      after: v2,
    });

    const brief = renderAgentBrief(
      v2,
      decideGovernance(await ledger.intents(), { strictTextSubmission: true }),
      { intents: await ledger.intents(), entries: await ledger.history() },
      { maxTargets: 12 },
    );

    say('## 6. Agent 判断依据（真实渲染）');
    say();
    say(`\`\`\`text\n${brief}\n\`\`\``);
    say();

    expect(brief).toContain('mode = strict');
    expect(brief).toContain('id=');
    expect(brief).toContain('H1 Quarterly Report');
    expect(brief).toContain('把营收改成 12%');
    expect(brief).toContain('Revenue grew by 10 percent. → Revenue grew by 12 percent.');
    expect(brief).toContain('added');
    // 定位纪律必须传达给 Agent。
    expect(brief).toContain('不要用「第 N 段」');
  });

  it('同一输入两次渲染逐字节一致', async () => {
    const v2 = await executeIR('rev-v2.docx');
    const decision = decideGovernance([]);
    const ledgerLike = { intents: [], entries: [] };
    expect(renderAgentBrief(v2, decision, ledgerLike)).toBe(renderAgentBrief(v2, decision, ledgerLike));
  });
});

describe('blackbox: 双 IR 映射（对应哈希表的三个方向）', () => {
  it('三条查找方向都能在真实文档上互指回到同一个节点', async () => {
    const content = await executeIR('rev-v1.docx');
    const map = content.sourceMap;

    say('## 7. 双 IR 映射');
    say();
    say('对应表提供三个方向的查找，缺一不可：');
    say();
    say('| 方向 | 输入 | 输出 | 用途 |');
    say('| --- | --- | --- | --- |');
    say('| 正向 | semanticId | 物理指针 | 「要改这个节点，动 XML 哪里」 |');
    say('| 反向 | 物理指针 | semanticId | 「这个 XML 元素是什么」 |');
    say('| 内容寻址 | 文本指纹 | semanticId 列表 | 「用户说的那句话在哪」 |');
    say();

    say('以段落「Revenue grew by 10 percent.」为例：');
    say();
    const target = content.semantic.blocks.find(
      (block) => block.kind === 'paragraph' && block.text === 'Revenue grew by 10 percent.',
    );
    expect(target).toBeDefined();
    if (target === undefined) return;

    // 方向一：知道语义节点，找到物理落点。
    const pointers = map.bySemanticId[target.id];
    expect(pointers).toEqual([`word/document.xml!${target.anchor.structuralPath}`]);
    say(`- 正向：semanticId ${target.id.slice(0, 12)}… → ${pointers?.[0]}`);

    // 方向二：拿着物理指针，问它对应哪个语义节点。
    const pointer = pointers?.[0] ?? '';
    expect(map.byPointer[pointer]).toBe(target.id);
    say(`- 反向：${pointer} → semanticId ${map.byPointer[pointer]?.slice(0, 12)}…`);

    // 方向三：只拿一段文字（用户原话里摘的），找回节点。
    // 注意 kind 用的是 OOXML 本地名（段落是 'p'，表格是 'tbl'），
    // 不是语义视图里的 'paragraph' / 'table'——指纹的输入是物理层事实。
    const digest = computeNodeDigest(target.anchor.kind, 'Revenue grew by 10 percent.');
    expect(map.byFingerprint[digest]).toContain(target.id);
    say(`- 内容寻址：指纹 ${digest.slice(0, 12)}… → semanticId ${map.byFingerprint[digest]?.[0]?.slice(0, 12)}…`);
    say();
    say('三条路径都落在同一个 semanticId 上，且经过 verify 的一致性检查。');
    say();

    // 三个索引必须两两自洽：这就是「对应哈希表」这三个字的分量。
    for (const entry of map.entries) {
      expect(map.bySemanticId[entry.id]).toContain(entry.pointer);
      expect(map.byPointer[entry.pointer]).toBe(entry.id);
      expect(map.byFingerprint[entry.anchor.digest]).toContain(entry.id);
    }
    expect(map.entries.length).toBeGreaterThan(0);
  });

  it('语义视图与物理视图逐节点一一对应，没有孤儿', async () => {
    const content = await executeIR('rev-v1.docx');
    const mappedPointers = content.sourceMap.entries.map((entry) => entry.pointer).sort();
    const physicalPointers = content.physical.nodes
      .map((node) => node.pointer)
      .filter((pointer) => mappedPointers.includes(pointer))
      .sort();

    say('### 两个视图是否真的对得上');
    say();
    say(`- 对应表登记 ${content.sourceMap.entries.length} 个语义节点`);
    say(`- 物理视图登记 ${content.physical.nodes.length} 个元素`);
    say(`- 其中 ${mappedPointers.length} 个物理元素被对应表指到`);
    say();

    expect(mappedPointers).toEqual(physicalPointers);
    // 语义节点数 == 被映射的物理元素数：没有任何一侧是孤立的。
    expect(mappedPointers).toHaveLength(content.sourceMap.entries.length);
  });

  it('节点位置漂移后，仍能靠内容把它找回来（冗余取证）', async () => {
    // 无 paraId 时 id 会随位置变，这是身份层的固有局限。
    // 但锚点同时带了内容选择器，所以「找不回 id」不等于「找不回节点」。
    const v1 = await executeIR('rev-drift-v1.docx');
    const v2 = await executeIR('rev-drift-v2.docx');

    const gammaV1 = v1.semantic.blocks.find(
      (block) => block.kind === 'paragraph' && block.text === 'Gamma',
    );
    expect(gammaV1).toBeDefined();
    if (gammaV1 === undefined) return;

    say('### 位置漂移后的重新定位');
    say();
    say('v1 与 v2 之间插入了一段，"Gamma" 从 w:p[3] 被顶到 w:p[4]。');
    say();
    say(`- v1 中 Gamma 的 id：${gammaV1.id.slice(0, 12)}…，位置 ${gammaV1.anchor.structuralPath}`);
    say(`- v1 与 v2 的 id 是否相同：${gammaV1.id === (v2.sourceMap.byFingerprint[gammaV1.anchor.digest]?.[0] ?? '')}`);
    say('  （没有 paraId，id 由位置派生，位置变了 id 必然变）');
    say();

    // 用 v1 取出的内容指纹，去 v2 的索引里反查。
    const relocatedIds = v2.sourceMap.byFingerprint[gammaV1.anchor.digest];
    expect(relocatedIds).toBeDefined();
    const relocated = v2.sourceMap.entries.find((entry) => entry.id === relocatedIds?.[0]);

    say(`- 但用 v1 的内容指纹去 v2 查，找回了：${relocated?.anchor.structuralPath}`);
    say();
    say('这就是锚点同时记三种选择器的意义：id 变了不代表节点失联，');
    say('内容选择器可以把它重新接上。**这正是将来写入模块敢动手的底气。**');
    say();

    expect(relocated?.anchor.structuralPath).toBe('/w:document/w:body/w:p[4]');
    expect(relocated?.anchor.quote).toBe('Gamma');
  });
});

describe('blackbox: 规模与上下文边界', () => {
  it('长文档的简报被截断，不会把上下文撑爆', async () => {
    const v2 = await executeIR('rev-v2.docx');
    const tiny = renderAgentBrief(v2, decideGovernance([]), { intents: [], entries: [] }, {
      maxTargets: 2,
    });
    const full = renderAgentBrief(v2, decideGovernance([]), { intents: [], entries: [] });

    say('## 8. 简报截断');
    say();
    say(`- maxTargets=2 时长度 ${tiny.length} 字符，含截断提示：${tiny.includes('未列出')}`);
    say(`- 默认输出长度 ${full.length} 字符`);
    say();

    expect(tiny).toContain('未列出');
    expect(tiny.length).toBeLessThan(full.length);
  });
});

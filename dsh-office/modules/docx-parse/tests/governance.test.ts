/**
 * 门槛判定与判断依据测试。
 *
 * 门槛判定是硬编码的业务规则，因此测试的重点不是「它聪明」，
 * 而是「它每次对同一个情形给出同一个结论，且结论能解释」。
 * 一旦有人改了规则，这里会立刻告警哪一条被动了。
 */
import { describe, expect, it } from 'vitest';

import {
  FORMAT_REVISION_INTENTION_THRESHOLD,
  createRevisionLedger,
  decideGovernance,
  deriveGovernanceSignals,
  diffDualIR,
  governanceDirective,
  renderAgentBrief,
  resolveGovernanceMode,
} from '../src/index';
import type { GovernanceSignals } from '../src/governance/contract';
import { buildIR, intent, para } from './support';

/** 造一组信号，只覆盖关心的开关。 */
function signals(overrides: Partial<GovernanceSignals> = {}): GovernanceSignals {
  return {
    explicitFormatRequirements: false,
    strictTextSubmission: false,
    strongFormatEditIntention: false,
    singleShotGeneration: false,
    ...overrides,
  };
}

/** 造 n 条 stage=revision、kind=format 的意图。 */
function formatRevisions(count: number) {
  return Array.from({ length: count }, (_, index) =>
    intent({ id: `f${index}`, stage: 'revision', kind: 'format', text: `字体再大一点(${index})` }),
  );
}

describe('governance: hard-coded decision table', () => {
  it('defaults to lite when no strict signal is present', () => {
    const decision = resolveGovernanceMode(signals());
    expect(decision.mode).toBe('lite');
    expect(decision.reason).toBe('no-strict-signal');
    expect(decision.matchedSignals).toEqual([]);
    expect(decision.exempted).toBe(false);
  });

  it('goes strict when the user stated format requirements up front', () => {
    const decision = resolveGovernanceMode(signals({ explicitFormatRequirements: true }));
    expect(decision.mode).toBe('strict');
    expect(decision.reason).toBe('explicit-format-requirements');
  });

  it('goes strict for a strict-text-submission scenario', () => {
    const decision = resolveGovernanceMode(signals({ strictTextSubmission: true }));
    expect(decision.mode).toBe('strict');
    expect(decision.reason).toBe('strict-text-submission');
  });

  it('goes strict on a strong format-edit intention', () => {
    const decision = resolveGovernanceMode(signals({ strongFormatEditIntention: true }));
    expect(decision.mode).toBe('strict');
    expect(decision.reason).toBe('format-edit-intention');
  });

  it('stays lite for a single-shot task with no other signal', () => {
    // 「帮我写个周报」：没有任何格式/文本严格信号 → 不该套严格治理。
    const decision = resolveGovernanceMode(signals({ singleShotGeneration: true }));
    expect(decision.mode).toBe('lite');
  });

  it('exempts a single-shot task that only drew format edits', () => {
    const decision = resolveGovernanceMode(
      signals({ singleShotGeneration: true, strongFormatEditIntention: true }),
    );
    expect(decision.mode).toBe('exempt');
    expect(decision.reason).toBe('single-shot-exempt');
    expect(decision.exempted).toBe(true);
  });

  it('does not exempt a single-shot task when format requirements were stated up front', () => {
    // 用户提前把格式说清楚了，说明他自己在意 —— 豁免让位。
    const decision = resolveGovernanceMode(
      signals({ singleShotGeneration: true, explicitFormatRequirements: true }),
    );
    expect(decision.mode).toBe('strict');
    expect(decision.exempted).toBe(false);
  });

  it('documents the deliberate edge: single-shot + strict scenario is exempt', () => {
    // 刻意的规则：单次生成是「这次请求」的事实，比宿主的场景粗判更贴近当下。
    // 这条断言的存在就是为了把该行为钉住，改规则时必然看到它。
    const decision = resolveGovernanceMode(
      signals({ singleShotGeneration: true, strictTextSubmission: true }),
    );
    expect(decision.mode).toBe('exempt');
  });

  it('reports every matched strict signal for explainability', () => {
    const decision = resolveGovernanceMode(
      signals({ explicitFormatRequirements: true, strictTextSubmission: true }),
    );
    expect(decision.matchedSignals).toEqual(['explicitFormatRequirements', 'strictTextSubmission']);
    // 理由取最高优先级那一条。
    expect(decision.reason).toBe('explicit-format-requirements');
  });

  it('is a pure function', () => {
    const input = signals({ strictTextSubmission: true });
    expect(resolveGovernanceMode(input)).toEqual(resolveGovernanceMode(input));
  });
});

describe('governance: deriving signals from the ledger', () => {
  it('treats a requirement-stage format intent as explicit requirements', () => {
    const intents = [intent({ stage: 'requirement', kind: 'format', text: '全篇黑体三号' })];
    expect(deriveGovernanceSignals(intents).explicitFormatRequirements).toBe(true);
  });

  it('does not treat a revision-stage format intent as a requirement', () => {
    const derived = deriveGovernanceSignals([intent({ stage: 'revision', kind: 'format' })]);
    expect(derived.explicitFormatRequirements).toBe(false);
  });

  it('needs repeated format revisions before calling the intention strong', () => {
    expect(deriveGovernanceSignals(formatRevisions(1)).strongFormatEditIntention).toBe(false);
    expect(
      deriveGovernanceSignals(formatRevisions(FORMAT_REVISION_INTENTION_THRESHOLD))
        .strongFormatEditIntention,
    ).toBe(true);
  });

  it('ignores content-only revisions when counting format intention', () => {
    const intents = Array.from({ length: 5 }, (_, index) =>
      intent({ id: `c${index}`, stage: 'revision', kind: 'content' }),
    );
    expect(deriveGovernanceSignals(intents).strongFormatEditIntention).toBe(false);
  });

  it('takes scenario signals only from the host context', () => {
    const derived = deriveGovernanceSignals([], { strictTextSubmission: true, singleShot: true });
    expect(derived.strictTextSubmission).toBe(true);
    expect(derived.singleShotGeneration).toBe(true);
  });

  it('accumulates evidence instead of letting one side silence the other', () => {
    // 账本说「有」、宿主说「没有」→ 仍然是有。沉默不等于否认。
    const derived = deriveGovernanceSignals(
      [intent({ stage: 'requirement', kind: 'format' })],
      { explicitFormatRequirements: false },
    );
    expect(derived.explicitFormatRequirements).toBe(true);
  });

  it('routes a single-shot request with an upfront requirement to strict', () => {
    const decision = decideGovernance(
      [intent({ stage: 'requirement', kind: 'format', text: '按模板来' })],
      { singleShot: true },
    );
    expect(decision.mode).toBe('strict');
  });

  it('stays lite for a single-shot request with no format signals at all', () => {
    // 没有任何严格信号 → 默认宽松，连豁免都谈不上。
    expect(decideGovernance([], { singleShot: true }).mode).toBe('lite');
  });

  it('goes strict after enough format revisions in a non-single-shot task', () => {
    const decision = decideGovernance(formatRevisions(FORMAT_REVISION_INTENTION_THRESHOLD));
    expect(decision.mode).toBe('strict');
    expect(decision.reason).toBe('format-edit-intention');
  });
});

describe('governance: directive text', () => {
  it('gives distinct, actionable instructions per mode', () => {
    const strict = governanceDirective('strict');
    const lite = governanceDirective('lite');
    const exempt = governanceDirective('exempt');

    expect(strict).toContain('semanticId');
    expect(strict).toContain('最小 diff');
    expect(exempt).toContain('单次生成');
    expect(lite).toContain('常规强度');
    // 三档文案必须互不相同，否则档位区分没有意义。
    expect(new Set([strict, lite, exempt]).size).toBe(3);
  });
});

describe('brief: agent decision basis', () => {
  const content = buildIR([
    para('Chapter One', 0, { outlineLevel: 0, headingStyle: true, paraId: 'H0000001' }),
    para('Some body text', 1),
  ]);

  it('includes the mode, the outline, addressable targets and the ledger', async () => {
    const ledger = createRevisionLedger();
    await ledger.record({ intent: intent({ id: 'i1' }) });
    await ledger.record({
      intent: intent({ id: 'i2', text: '把正文改短' }),
      after: buildIR([para('Chapter One', 0), para('Shorter text', 1)]),
    });

    const brief = renderAgentBrief(content, decideGovernance([], { strictTextSubmission: true }), {
      intents: await ledger.intents(),
      entries: await ledger.history(),
    });

    expect(brief).toContain('## 治理档位');
    expect(brief).toContain('mode = strict');
    expect(brief).toContain('## 文档大纲');
    expect(brief).toContain('## 可点名目标');
    expect(brief).toContain('## 修订账本');
    // 可点名目标必须带可复制的 id，否则 Agent 还是只能靠「第几段」。
    expect(brief).toContain('id=');
    expect(brief).toContain('H1 Chapter One');
    // 用户原话要原样出现，不能被改写。
    expect(brief).toContain('把正文改短');
    // 定位纪律必须出现在简报里。
    expect(brief).toContain('第 N 段');
  });

  it('renders deterministically', () => {
    const decision = decideGovernance([]);
    const ledgerLike = { intents: [], entries: [] };
    expect(renderAgentBrief(content, decision, ledgerLike)).toBe(
      renderAgentBrief(content, decision, ledgerLike),
    );
  });

  it('truncates long documents instead of blowing up the context', () => {
    const many = buildIR(
      Array.from({ length: 60 }, (_, index) => para(`paragraph number ${index}`, index)),
    );
    const brief = renderAgentBrief(many, decideGovernance([]), { intents: [], entries: [] }, {
      maxTargets: 5,
    });
    expect(brief).toContain('另有 55 个节点未列出');
  });

  it('says so plainly when there is nothing to point at', () => {
    const brief = renderAgentBrief(buildIR([]), decideGovernance([]), {
      intents: [],
      entries: [],
    });
    expect(brief).toContain('文档为空');
    expect(brief).toContain('没有标题层级');
    expect(brief).toContain('尚无修订记录');
  });

  it('flags inferred heading levels so the agent knows the level may be wrong', () => {
    const inferred = buildIR([para('Odd heading', 0, { headingStyle: true })]);
    const brief = renderAgentBrief(inferred, decideGovernance([]), { intents: [], entries: [] });
    expect(brief).toContain('（级别为推断）');
  });
});

describe('brief: end-to-end with a real diff', () => {
  it('shows the actual change computed by the identity-aware diff', async () => {
    const ledger = createRevisionLedger();
    const v1 = buildIR([para('Alpha', 0, { paraId: 'A1' }), para('Gamma', 1, { paraId: 'C1' })]);
    const v2 = buildIR([
      para('Alpha', 0, { paraId: 'A1' }),
      para('Beta', 1, { paraId: 'B1' }),
      para('Gamma', 2, { paraId: 'C1' }),
    ]);

    await ledger.record({ intent: intent({ id: 'i1' }), after: v1 });
    await ledger.record({ intent: intent({ id: 'i2', text: '中间加一段' }), after: v2 });

    const brief = renderAgentBrief(v2, decideGovernance([], { strictTextSubmission: true }), {
      intents: await ledger.intents(),
      entries: await ledger.history(),
    });

    // 简报里应能看出「只新增了一段」，而不是「后面全变了」。
    expect(brief).toContain('added');
    expect(brief).toContain('Beta');
    expect(diffDualIR(v1, v2)).toHaveLength(1);
  });
});

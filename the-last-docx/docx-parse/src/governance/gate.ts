/**
 * 门槛判定 —— 一个刻意【硬编码】的条件。
 *
 * 为什么硬编码而不做成配置项：
 *   这是业务意图，不是可调参数。做成配置会让每个调用方各自拍一个值，
 *   最后没人说得清「什么情况该严格」。先把唯一一份判断写死在代码里，
 *   等真实使用中出现扛不住的场景，再带着证据去改——而不是提前给一堆开关。
 *
 * 这一层将来会整体搬成一个独立模块（`governance/` 目录就是为此划分的边界），
 * 因此本文件不依赖 docx-parse 的任何实现，只依赖 `revision/contract.ts` 的类型。
 *
 * --- 硬编码的判定规则（自上而下，命中即返回）---
 *
 *   1. 三个严格信号一个都没命中            → lite     （默认宽松）
 *   2. 单次生成，且用户没先给格式要求       → exempt   （豁免）
 *   3. 用户先给了格式要求                  → strict
 *   4. 场景对文本提交有严格要求             → strict
 *   5. 用户格式修改意愿明显                → strict
 *
 * 第 2 条优先于第 3~5 条，但【让位】于「先给了格式要求」：
 * 用户提前把格式要求说清楚，说明他自己在意，此时不该拿「单次生成」当理由放松。
 * 反过来，「单次生成 + 场景严格」这一组合会走豁免——这是刻意的：
 * 场景标签是宿主的粗判，而单次生成是具体任务的事实，后者更贴近当下这次请求。
 * 若实际使用中发现该组合不该豁免，改这一处即可。
 */
import type { RevisionIntent } from '../revision/contract';
import type {
  GovernanceContext,
  GovernanceDecision,
  GovernanceMode,
  GovernanceReason,
  GovernanceSignals,
  StrictSignalName,
} from './contract';

/**
 * 判定「格式修改意愿明显」所需的格式类修改意图条数。
 *
 * 为什么是 2 而不是 1：单次提及很可能只是顺口一说（「顺便字体换一下」），
 * 连续多次才说明用户是真的在意格式。这个数字是硬编码的经验值，
 * 唯一的意义是让它有一个【有名字的】落点，将来调它时知道在调什么。
 */
export const FORMAT_REVISION_INTENTION_THRESHOLD = 2;

/** 从意图列表里提取与门槛判定相关的计数。 */
function countIntentEvidence(intents: readonly RevisionIntent[]): {
  hasRequirementFormat: boolean;
  formatRevisionCount: number;
} {
  let hasRequirementFormat = false;
  let formatRevisionCount = 0;

  for (const intent of intents) {
    if (intent.kind !== 'format') continue;
    if (intent.stage === 'requirement') {
      hasRequirementFormat = true;
    } else {
      formatRevisionCount += 1;
    }
  }
  return { hasRequirementFormat, formatRevisionCount };
}

/**
 * 综合账本与宿主上下文，得到完整的判定信号。
 *
 * 证据只会累加：账本说「有」、宿主说「有」，结果就是「有」；
 * 任何一方的肯定都不会被另一方的沉默抹掉。沉默不代表否认——
 * 那只是「没看到」，而本层的默认是宽松，所以沉默自然导向宽松。
 */
export function deriveGovernanceSignals(
  intents: readonly RevisionIntent[],
  context: GovernanceContext = {},
): GovernanceSignals {
  const evidence = countIntentEvidence(intents);

  return {
    // 账本推断与宿主声明取或。
    explicitFormatRequirements:
      evidence.hasRequirementFormat || context.explicitFormatRequirements === true,
    strictTextSubmission: context.strictTextSubmission === true,
    strongFormatEditIntention:
      evidence.formatRevisionCount >= FORMAT_REVISION_INTENTION_THRESHOLD,
    singleShotGeneration: context.singleShot === true,
  };
}

/** 按固定顺序列出命中的严格信号（顺序稳定，便于断言与展示）。 */
function collectStrictSignals(signals: GovernanceSignals): StrictSignalName[] {
  const matched: StrictSignalName[] = [];
  if (signals.explicitFormatRequirements) matched.push('explicitFormatRequirements');
  if (signals.strictTextSubmission) matched.push('strictTextSubmission');
  if (signals.strongFormatEditIntention) matched.push('strongFormatEditIntention');
  return matched;
}

/** 有严格信号时，按优先级给出具体结论码。 */
function strictReason(matched: readonly StrictSignalName[]): GovernanceReason {
  if (matched.includes('explicitFormatRequirements')) return 'explicit-format-requirements';
  if (matched.includes('strictTextSubmission')) return 'strict-text-submission';
  return 'format-edit-intention';
}

/**
 * 由信号得出治理强度。
 *
 * 纯函数：同样输入永远同样输出，不读时间、不读配置、无副作用。
 */
export function resolveGovernanceMode(signals: GovernanceSignals): GovernanceDecision {
  const matchedSignals = collectStrictSignals(signals);

  // 规则 1：默认宽松。
  if (matchedSignals.length === 0) {
    return {
      mode: 'lite',
      reason: 'no-strict-signal',
      matchedSignals,
      exempted: false,
      signals,
    };
  }

  // 规则 2：单次生成豁免，但让位于「先给了格式要求」。
  if (signals.singleShotGeneration && !signals.explicitFormatRequirements) {
    return {
      mode: 'exempt',
      reason: 'single-shot-exempt',
      matchedSignals,
      exempted: true,
      signals,
    };
  }

  // 规则 3~5：走严格，并按优先级给出理由。
  return {
    mode: 'strict',
    reason: strictReason(matchedSignals),
    matchedSignals,
    exempted: false,
    signals,
  };
}

/**
 * 一步到位的便捷入口：给账本意图与宿主上下文，直接拿判定结论。
 *
 * 之所以保留两步（derive / resolve）也导出：调试时经常需要单独看
 * 「信号是怎么来的」与「结论是怎么下的」。
 */
export function decideGovernance(
  intents: readonly RevisionIntent[],
  context: GovernanceContext = {},
): GovernanceDecision {
  return resolveGovernanceMode(deriveGovernanceSignals(intents, context));
}

/**
 * 各档治理强度的执行纪律说明。
 *
 * 这是给 Agent 看的「判断依据」的一部分：光说「现在是 strict」没用，
 * 得说清楚 strict 意味着什么。文案按档位硬编码，避免每次生成时措辞漂移。
 */
export function governanceDirective(mode: GovernanceMode): string {
  switch (mode) {
    case 'strict':
      return [
        'STRICT: 本次改动必须走完整纪律。',
        '- 每次写入前先声明目标节点的 semanticId，禁止按顺序号或「第几段」指代。',
        '- 改完必须重算双 IR 并核对锚点仍然成立（对应表不得出现孤儿或抢占）。',
        '- 优先做最小 diff：只动目标节点，不重写整段/整表。',
        '- 不确定用户意图时先问，不要猜着改。',
      ].join('\n');
    case 'exempt':
      return [
        'EXEMPT: 单次生成类任务，允许直接给结果。',
        '- 不必逐节点声明目标，也不必强制最小 diff。',
        '- 但若用户随后提出修改，应立即升级为 STRICT，并把之前的产出登记进账本。',
      ].join('\n');
    default:
      return [
        'LITE: 常规强度。',
        '- 改动仍应基于 semanticId 定位，但不强制逐步核对。',
        '- 一旦用户开始提格式要求，立即升级为 STRICT。',
      ].join('\n');
  }
}

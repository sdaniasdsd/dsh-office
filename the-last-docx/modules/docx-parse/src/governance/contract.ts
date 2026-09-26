/**
 * 门槛判定层契约。
 *
 * 存在的理由（用户原话）：并不是所有办公文件都要求如此严格。
 * 「单次帮我写个周报」和「按这份公文格式要求改」是两种截然不同的任务，
 * 用同一套严格治理去套前者，只会让 Agent 变慢变啰嗦。
 *
 * 因此本层的唯一职责是：判断当前该用哪一档治理强度。
 * 它【不】执行治理，只给出结论与理由，让调用方自己决定怎么用。
 */

/** 治理强度。 */
export const GOVERNANCE_MODES = ['strict', 'lite', 'exempt'] as const;
export type GovernanceMode = (typeof GOVERNANCE_MODES)[number];

/**
 * 判定所依据的信号。
 *
 * 四个信号刻意分开而不是揉成一个分数：
 * 分数无法解释「为什么这么判」，而本层的结论要能对用户讲清楚。
 */
export interface GovernanceSignals {
  /**
   * 用户动手【之前】就给了格式要求。
   * 例如「按公司模板，标题黑体三号」——这是最硬的约束信号。
   */
  explicitFormatRequirements: boolean;
  /**
   * 场景本身对文本提交有严格要求（合同、公文、标书、法律文书等）。
   * 这是【场景属性】，通常由宿主按业务线声明，无法从对话里推断。
   */
  strictTextSubmission: boolean;
  /**
   * 用户的格式修改意愿明显。
   * 依据是「修改阶段反复提到格式」——一次可能是顺口一提，多次才算意愿。
   */
  strongFormatEditIntention: boolean;
  /**
   * 办公单次生成类任务。
   * 例如「帮我写个会议纪要」——产出即交付，没有反复打磨的过程。
   */
  singleShotGeneration: boolean;
}

/** 命中严格信号的具体条目，用于向用户解释判定理由。 */
export type StrictSignalName =
  | 'explicitFormatRequirements'
  | 'strictTextSubmission'
  | 'strongFormatEditIntention';

/** 判定结论。 */
export interface GovernanceDecision {
  mode: GovernanceMode;
  /** 稳定、可编程判断的结论码（便于日志与回归断言）。 */
  reason: GovernanceReason;
  /** 命中的严格信号；mode 为 lite 时为空数组。 */
  matchedSignals: StrictSignalName[];
  /** 是否因「单次生成」而获得豁免。 */
  exempted: boolean;
  /** 判定所依据的全部信号，原样回填，便于排查。 */
  signals: GovernanceSignals;
}

/** 结论码。 */
export const GOVERNANCE_REASONS = [
  /** 无任何严格信号 → 默认宽松。 */
  'no-strict-signal',
  /** 用户先给了格式要求 → 严格。 */
  'explicit-format-requirements',
  /** 场景对文本提交有严格要求 → 严格。 */
  'strict-text-submission',
  /** 用户格式修改意愿明显 → 严格。 */
  'format-edit-intention',
  /** 单次生成类且未先给格式要求 → 豁免。 */
  'single-shot-exempt',
] as const;
export type GovernanceReason = (typeof GOVERNANCE_REASONS)[number];

/**
 * 宿主提供的上下文。
 *
 * 与「从账本推断」互补：账本能看出用户说了什么，
 * 但「这是不是合同场景」「是不是单次任务」只有宿主知道。
 */
export interface GovernanceContext {
  /** 本次是否为单次生成类任务。 */
  singleShot?: boolean;
  /** 场景是否对文本提交有严格要求。 */
  strictTextSubmission?: boolean;
  /**
   * 宿主已自行判定「用户是否先给了格式要求」。
   * 提供了就以它为准并与账本推断取或——证据只会累加，不会互相抵消。
   */
  explicitFormatRequirements?: boolean;
}

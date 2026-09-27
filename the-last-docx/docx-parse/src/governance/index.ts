/**
 * 门槛判定层出口。
 *
 * 这就是用户说的「到时候我要新起一个模块」的那块：本目录已经按独立模块
 * 的边界划分——只依赖 `revision/contract.ts` 的类型，不碰 docx-parse 的任何实现。
 * 新起模块时把这个目录整体搬走即可。
 */
export { GOVERNANCE_MODES, GOVERNANCE_REASONS } from './contract';
export type {
  GovernanceContext,
  GovernanceDecision,
  GovernanceMode,
  GovernanceReason,
  GovernanceSignals,
  StrictSignalName,
} from './contract';

export {
  FORMAT_REVISION_INTENTION_THRESHOLD,
  decideGovernance,
  deriveGovernanceSignals,
  governanceDirective,
  resolveGovernanceMode,
} from './gate';

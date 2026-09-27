/**
 * 修订层出口。
 *
 * 这个 barrel 是「将来搬成独立模块」的接缝：搬走时把这个目录整体移走，
 * 再把下面的 `../contract` 引用换成新模块自己的类型即可，其余代码不动。
 */
export {
  CHANGE_FIELDS,
  INTENT_KINDS,
  INTENT_SOURCES,
  INTENT_STAGES,
} from './contract';
export type {
  Awaitable,
  ChangeField,
  IntentKind,
  IntentSource,
  IntentStage,
  RecordRevisionInput,
  RevisionChange,
  RevisionEntry,
  RevisionIntent,
  RevisionLedger,
  RevisionStore,
} from './contract';

export { collectFacets, diffDualIR, irFingerprint } from './diff';
export type { DiffOptions } from './diff';

export { createMemoryRevisionStore, createRevisionLedger, validateIntent } from './ledger';
export type { CreateLedgerOptions } from './ledger';

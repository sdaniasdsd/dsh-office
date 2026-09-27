/**
 * 修订账本 —— 接收外部意图并维护修订历史。
 *
 * 数据流：
 *
 *   外部接口 → RevisionIntent（原话 + 标签）
 *            → validateIntent（信任边界）
 *            → 与上一版双 IR 做差（diff.ts）
 *            → RevisionEntry 追加进 store
 *            → store 更新基线（供下次做差）
 *
 * 关于「实时」：本层不引入事件循环或订阅表，而是通过 `onRecord` 回调把
 * 每次记账同步推给宿主。宿主若需要扇出，用自己的事件总线接这个回调即可——
 * 这样本层保持零依赖、可单独测试。
 */
import { DocxParseError } from '../errors';
import type { DocxDualIR, SemanticId } from '../contract';
import { INTENT_KINDS, INTENT_SOURCES, INTENT_STAGES } from './contract';
import type {
  RecordRevisionInput,
  RevisionEntry,
  RevisionIntent,
  RevisionLedger,
  RevisionStore,
} from './contract';
import { diffDualIR, irFingerprint } from './diff';

/** 把任意未知值断言为字符串，否则抛出 INVALID_INPUT。 */
function asString(value: unknown, field: string): string {
  if (typeof value !== 'string') {
    throw new DocxParseError('INVALID_INPUT', `intent.${field} must be a string`, {
      details: { field, received: typeof value },
    });
  }
  return value;
}

/** 断言值属于给定枚举。 */
function asEnum<T extends string>(value: unknown, allowed: readonly T[], field: string): T {
  if (typeof value !== 'string' || !(allowed as readonly string[]).includes(value)) {
    throw new DocxParseError('INVALID_INPUT', `intent.${field} must be one of ${allowed.join('|')}`, {
      details: { field, received: typeof value === 'string' ? value : typeof value },
    });
  }
  return value as T;
}

/**
 * 校验外部传入的意图是否符合契约。
 *
 * 意图来自模块外部（对话入口、表单、上游解析器），因此必须当作【不可信输入】：
 * 逐字段断言后才允许进入账本。否则一条脏数据会污染整条修订历史，
 * 而修订历史是只追加的——脏数据一旦进去就再也擦不掉。
 */
export function validateIntent(value: unknown): RevisionIntent {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new DocxParseError('INVALID_INPUT', 'intent must be a plain object');
  }
  const record = value as Record<string, unknown>;

  const id = asString(record['id'], 'id');
  if (id.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'intent.id must be non-empty');
  }
  const requestId = asString(record['requestId'], 'requestId');
  const text = asString(record['text'], 'text');
  if (text.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'intent.text must be non-empty');
  }

  const rawTargets = record['targetIds'];
  if (!Array.isArray(rawTargets)) {
    throw new DocxParseError('INVALID_INPUT', 'intent.targetIds must be an array');
  }
  const targetIds: SemanticId[] = rawTargets.map((entry, index) =>
    asString(entry, `targetIds[${index}]`),
  );

  return {
    id,
    requestId,
    source: asEnum(record['source'], INTENT_SOURCES, 'source'),
    stage: asEnum(record['stage'], INTENT_STAGES, 'stage'),
    kind: asEnum(record['kind'], INTENT_KINDS, 'kind'),
    text,
    targetIds,
  };
}

/** 创建模块时可选的注入点。 */
export interface CreateLedgerOptions {
  /** 时间源，默认 `Date.now`。注入后测试可得到完全确定的时间戳。 */
  now?: () => number;
  /** 每次记账后的同步回调，宿主据此实现「实时」推送。 */
  onRecord?: (entry: RevisionEntry) => void;
}

/**
 * 内存存储实现。
 *
 * 用途：单测、单进程独立运行，以及「宿主暂时没有持久化需求」的场景。
 * 它保存的是引用而非深拷贝——账本约定调用方不修改已交付的 IR。
 */
export function createMemoryRevisionStore(): RevisionStore {
  const entries: RevisionEntry[] = [];
  let baseline: DocxDualIR | null = null;

  return {
    loadEntries(): RevisionEntry[] {
      // 返回浅拷贝数组，防止外部 push 破坏只追加语义。
      return [...entries];
    },
    appendEntry(entry: RevisionEntry): void {
      entries.push(entry);
    },
    loadBaseline(): DocxDualIR | null {
      return baseline;
    },
    saveBaseline(content: DocxDualIR): void {
      baseline = content;
    },
  };
}

/**
 * 构建修订账本。
 *
 * 纯函数式工厂：不注册全局副作用，可创建多个互不干扰的实例。
 */
export function createRevisionLedger(
  store: RevisionStore = createMemoryRevisionStore(),
  options: CreateLedgerOptions = {},
): RevisionLedger {
  const now = options.now ?? (() => Date.now());

  /** 按修订号升序读取全部记录。 */
  const loadHistory = async (): Promise<RevisionEntry[]> => {
    const entries = await store.loadEntries();
    return [...entries].sort((left, right) => left.revision - right.revision);
  };

  return {
    async record(input: RecordRevisionInput): Promise<RevisionEntry> {
      // 信任边界：外部意图先校验，再落账。
      const intent = validateIntent(input.intent);

      const existing = await store.loadEntries();
      // 修订号单调递增：取现有最大值 + 1，避免删除后回退（只追加本就不会回退，
      // 但这里显式算最大值，即使 store 实现有缺口也不会重号）。
      const nextRevision = existing.reduce((max, entry) => Math.max(max, entry.revision), 0) + 1;

      const baseline = await store.loadBaseline();
      const producedVersion = input.after !== undefined;
      // 未产生新版本（或首条记录无基线）时如实记为空差异，不伪造「全量新增」。
      const changes =
        input.after !== undefined && baseline !== null ? diffDualIR(baseline, input.after) : [];

      const entry: RevisionEntry = {
        revision: nextRevision,
        intent,
        recordedAtMs: now(),
        changes,
        beforeFingerprint: baseline === null ? null : irFingerprint(baseline),
        afterFingerprint: input.after === undefined ? null : irFingerprint(input.after),
        producedVersion,
      };

      await store.appendEntry(entry);
      // 基线后置更新：先落账再更新，确保账本与基线不会出现「账记了但基线没跟上」
      // 之外的中间态。
      if (input.after !== undefined) {
        await store.saveBaseline(input.after);
      }
      options.onRecord?.(entry);
      return entry;
    },

    async history(): Promise<RevisionEntry[]> {
      return loadHistory();
    },

    async intents(): Promise<RevisionIntent[]> {
      const entries = await loadHistory();
      return entries.map((entry) => entry.intent);
    },
  };
}

/**
 * 修订层契约 —— 外部接口的形状定义。
 *
 * 这一层要解决的是「改动缺少中介机制，全凭大模型本身」这个缺口：
 *
 *   双 IR 回答了「文档现在长什么样、每个节点是谁」；
 *   本层回答「用户想要什么、已经改过什么、这次改动动了哪里」。
 *
 * 三者合起来，Agent 才有一个【可查证】的判断依据，而不是只靠对话上下文即兴发挥。
 *
 * 两个刻意的设计决定：
 *   1. 意图用「原话 + 结构化标签」双轨记录。原话是证据，标签是索引；
 *      绝不把原话改写成标签后丢掉原话——那是不可逆的信息损失。
 *   2. 账本是【只追加】的。修订历史一旦写入就不再改写，
 *      这样「为什么现在是这个样子」永远可以回溯。
 */
import type { DocxDualIR, SemanticId } from '../contract';

/** 允许同步或异步的实现，便于宿主换成数据库而无需改函数签名。 */
export type Awaitable<T> = T | Promise<T>;

/**
 * 意图阶段 —— 区分「先给的要求」与「后提的意见」。
 *
 * 这个区分不是修辞：门槛判定要区分
 *   - requirement + format  → 用户动手前就给了格式要求（强约束信号）
 *   - revision  + format    → 用户看完结果想改格式（意愿信号）
 * 二者的治理含义完全不同。
 */
export const INTENT_STAGES = ['requirement', 'revision'] as const;
export type IntentStage = (typeof INTENT_STAGES)[number];

/**
 * 意图类型 —— 用户到底想动哪一类东西。
 *
 * `format` 是门槛判定最关心的那一类：它决定要不要走严格路径。
 * `unknown` 是诚实的兜底：上游没解析出来时不猜。
 */
export const INTENT_KINDS = ['format', 'content', 'structure', 'meta', 'unknown'] as const;
export type IntentKind = (typeof INTENT_KINDS)[number];

/** 意图来源。区分来源是为了在账本里分清「用户说的」与「Agent 推断的」。 */
export const INTENT_SOURCES = ['user', 'agent', 'system'] as const;
export type IntentSource = (typeof INTENT_SOURCES)[number];

/**
 * 一条用户意图（或 Agent 代为登记的意图）。
 *
 * `id` 由【宿主】提供而不由本层生成：生成 id 需要时间戳或随机数，
 * 会破坏本模块「同一输入产出完全一致输出」的确定性纪律。
 * 宿主本来就知道自己的请求身份（requestId / 消息 id），复用它即可。
 */
export interface RevisionIntent {
  /** 稳定标识，由宿主提供。 */
  id: string;
  /** 关联的请求 id，用于把意图与那次调用串起来。 */
  requestId: string;
  source: IntentSource;
  stage: IntentStage;
  kind: IntentKind;
  /** 用户原话（或 Agent 转述），保留原文不做改写。 */
  text: string;
  /**
   * 用户点名要改的语义节点 id。
   * 为空数组表示「全局意见」（例如「整体再简洁一点」），不对应具体节点。
   */
  targetIds: SemanticId[];
}

/**
 * 一处节点级改动。
 *
 * `fields` 是关键：它说明这次改动【动的是文字还是格式】。
 * 只比对文本指纹是不够的——把 H2 改成 H3，文字一个字没变，
 * 但那正是本层最需要看见的那类改动。
 */
export interface RevisionChange {
  semanticId: SemanticId;
  kind: string;
  change: 'added' | 'removed' | 'modified' | 'unchanged';
  /**
   * 变化的维度：
   *   text      = 正文内容变了；
   *   structure = 表格的行列结构变了（表格层只比结构，不比内容）；
   *   style     = 样式引用变了；
   *   level     = 标题级别变了。
   */
  fields: ChangeField[];
  /** 改动前的可读摘要（截断），null 表示该侧不存在。 */
  before: string | null;
  /** 改动后的可读摘要（截断）。 */
  after: string | null;
  /** 该节点对应的物理落点，写入端直接可用。 */
  pointers: string[];
}

export const CHANGE_FIELDS = ['text', 'structure', 'style', 'level'] as const;
export type ChangeField = (typeof CHANGE_FIELDS)[number];

/** 一条修订记录：一次意图 + 它带来的实际改动。 */
export interface RevisionEntry {
  /** 单调递增的修订号，从 1 开始。 */
  revision: number;
  intent: RevisionIntent;
  /** 记账时刻（毫秒）。 */
  recordedAtMs: number;
  /** 相对上一版的节点级改动；未产生新版本或首条记录时为空数组。 */
  changes: RevisionChange[];
  /** 改动前语义视图的指纹；无基线时为 null。 */
  beforeFingerprint: string | null;
  /** 改动后语义视图的指纹；本次未产生新版本时为 null。 */
  afterFingerprint: string | null;
  /** 本次是否产生了新的文档版本。 */
  producedVersion: boolean;
}

/**
 * 存储端口 —— 宿主可注入内存、文件或数据库实现。
 *
 * 为什么要 `loadBaseline` / `saveBaseline` 而不要求调用方每次传 `before`：
 * 让「上一版是什么」由账本自己维护，宿主就没有传错的机会。
 */
export interface RevisionStore {
  /** 按修订号升序读取全部记录。 */
  loadEntries(): Awaitable<RevisionEntry[]>;
  /** 追加一条记录。 */
  appendEntry(entry: RevisionEntry): Awaitable<void>;
  /** 读取上一次登记的双 IR 快照；从未登记过则为 null。 */
  loadBaseline(): Awaitable<DocxDualIR | null>;
  /** 覆盖保存最新的双 IR 快照。 */
  saveBaseline(content: DocxDualIR): Awaitable<void>;
}

/** 账本门面：宿主拿到的就是这四个方法。 */
export interface RevisionLedger {
  /**
   * 登记一条意图，并把「应用后的双 IR」与上一版做差。
   *
   * 这是「用户修改意见怼过来能收得到且能更新」的落点：
   * 收 → 校验 → 记账 → 算差 → 更新基线，一步到位。
   */
  record(input: RecordRevisionInput): Awaitable<RevisionEntry>;
  /** 全部修订记录（修订号升序）。 */
  history(): Awaitable<RevisionEntry[]>;
  /** 全部意图（记账顺序），供门槛判定使用。 */
  intents(): Awaitable<RevisionIntent[]>;
}

/** 登记一次修订所需的输入。 */
export interface RecordRevisionInput {
  intent: RevisionIntent;
  /**
   * 应用该意图【之后】的双 IR；省略表示这次只是「表达了意图」，文档尚未变化
   * （例如用户动手前先给的格式要求）。
   *
   * 为什么要求传「之后」而不是「改动本身」：改动应该由写入端真正落到 OOXML 上
   * 再重新解析出双 IR，账本记录的是【既成事实】，而不是「本打算改什么」。
   * 这样账本天然无法被伪造。
   */
  after?: DocxDualIR;
}

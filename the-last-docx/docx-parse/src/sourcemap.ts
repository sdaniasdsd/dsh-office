/**
 * 身份层 —— 双 IR 的「账本」，也是本模块相对只读探测模块最核心的增量。
 *
 * 为什么需要单独的层？
 *   引擎只陈述「看到的事实」，而「这个节点是谁」「它的稳定 id 是什么」
 *   「语义节点与物理节点如何对上号」必须由模块自己拥有。把身份逻辑集中在这里，
 *   好处有两个：
 *     1. 换引擎不会改变身份方案（id 的算法不依赖任何引擎特性）；
 *     2. 身份方案可被单独测试与验证（`verifier.ts` 会重算一遍来确认确定性）。
 *
 * 身份方案：composite-anchor-v1。核心是【冗余取证】——
 * 不赌任何单一选择器，而是同时记录三种彼此独立的定位依据：
 *
 *     | 选择器         | 独立性来源 | 失效场景      | 由谁兜底       |
 *     | -------------- | ---------- | ------------- | -------------- |
 *     | paraId         | 原生 id    | 可选、常缺失  | quote / path   |
 *     | quote          | 内容       | 内容被编辑    | paraId / path  |
 *     | structuralPath | 位置       | 增删导致漂移  | paraId / quote |
 *
 * 只要还有一条选择器有效，节点就能被重新找回；三条都失效才算真正「失联」。
 * 这正是 W3C Web Annotation 用「位置 + 引文」双选择器冗余标注同一目标的思路。
 */
import { createHash } from 'node:crypto';

import {
  SOURCE_MAP_SCHEME,
  type NodeAnchor,
  type SemanticId,
  type SourceMap,
  type SourceMapEntry,
  type SourceMapScheme,
} from './contract';

/** 引文选择器保留的最大字符数。过长会让 IR 变得臃肿，且截断后定位能力几乎不变。 */
const MAX_QUOTE_LENGTH = 160;

/** 语义 id 保留的摘要字符数（128 位，碰撞概率可忽略）。 */
const SEMANTIC_ID_LENGTH = 32;

/** 计算 SHA-256 十六进制摘要。集中在此，保证全模块摘要算法一致。 */
export function sha256Hex(input: string): string {
  return createHash('sha256').update(input, 'utf8').digest('hex');
}

/**
 * 归一化文本，作为内容选择器与指纹的输入。
 *
 * 规则：把**排版层面的空白**（空格、制表符、换行、不换行空格）折叠为单个空格，再去除首尾同类空白。
 * 之所以这样做：OOXML 会把同一句话拆到多个 run 里，中间夹带格式标记，
 * 若不归一化，「同一段文字」在两次解析间可能得到不同指纹。
 *
 * **不含 U+3000（表意空格）**。JavaScript 的 `\s` 同时匹配 U+3000 和 U+00A0，
 * 但 U+3000 在中文里是作者键入的可见字符（Word 按字符存储），不是排版空白。
 * 早先的实现把 `第一条　合同期限` 折叠为 `第一条 合同期限`，于是 IR 暴露的 `text`
 * 并不是文档里的文字，凡按 IR 文本回转的操作都会丢掉这个字符。
 */
const LAYOUT_SPACE = /[\t\n\r\f\v \u00a0]+/g;
const LAYOUT_EDGE = /^[\t\n\r\f\v \u00a0]+|[\t\n\r\f\v \u00a0]+$/g;
export function normalizeText(text: string): string {
  return text.replace(LAYOUT_SPACE, ' ').replace(LAYOUT_EDGE, '');
}

/**
 * 内容指纹：kind 与归一化文本的摘要。
 *
 * 用 \u0000 作为分隔符——它不可能出现在 OOXML 的本地名或正文里，
 * 因此不会出现「(ab, c) 与 (a, bc) 撞成同一指纹」这类拼接歧义。
 */
export function computeNodeDigest(kind: string, text: string): string {
  return sha256Hex(`${kind}\u0000${normalizeText(text)}`);
}

/** 构造引文选择器（截断后的归一化文本）。 */
export function buildQuote(text: string): string {
  const normalized = normalizeText(text);
  if (normalized.length <= MAX_QUOTE_LENGTH) return normalized;
  return normalized.slice(0, MAX_QUOTE_LENGTH);
}

/** 构造物理指针：`<part>!<structuralPath>`。 */
export function makePointer(part: string, structuralPath: string): string {
  return `${part}!${structuralPath}`;
}

/** 构造锚点所需的原始观察。 */
export interface AnchorInput {
  kind: string;
  part: string;
  structuralPath: string;
  ordinal: number;
  paraId: string | null;
  text: string;
}

/**
 * 构造一个节点锚点。
 *
 * 注意这里【不抛错】：哪怕所有选择器都很弱（空段落且无 paraId），
 * 也照样产出锚点，只是它的冗余度低。是否要就此告警由调用方决定
 * （见 `anchorSelectorCount` 与 mapper 的 ANCHOR_INCOMPLETE）。
 */
export function buildAnchor(input: AnchorInput): NodeAnchor {
  return {
    kind: input.kind,
    part: input.part,
    structuralPath: input.structuralPath,
    ordinal: input.ordinal,
    paraId: input.paraId,
    quote: buildQuote(input.text),
    digest: computeNodeDigest(input.kind, input.text),
  };
}

/**
 * 计算语义 id。
 *
 * 主键优先取 paraId（原生、稳定、不随位置漂移），否则退化为结构路径。
 * 之所以不把内容纳入 id：内容一改 id 就变，节点会被当成陌生人，
 * 挂在它身上的批注/书签就会失联。内容选择器的作用是「重新找回」，
 * 而不是「定义身份」——两者职责必须分开。
 */
export function computeSemanticId(anchor: NodeAnchor): SemanticId {
  const primaryKey = anchor.paraId !== null && anchor.paraId !== ''
    ? anchor.paraId
    : anchor.structuralPath;
  return sha256Hex(`${anchor.part}\u0000${anchor.kind}\u0000${primaryKey}`).slice(
    0,
    SEMANTIC_ID_LENGTH,
  );
}

/**
 * 统计锚点携带的【独立】选择器数量（1..3）。
 *
 * 三个族系：原生 id、内容、位置。注意 digest 由内容派生，因此与 quote 同族，
 * 不重复计数——否则会给一个实际上只有一种依据的锚点虚高的冗余度。
 */
export function anchorSelectorCount(anchor: NodeAnchor): number {
  let count = 0;
  if (anchor.paraId !== null && anchor.paraId !== '') count += 1;
  if (anchor.quote !== '') count += 1;
  if (anchor.structuralPath !== '') count += 1;
  return count;
}

/** 锚点冗余度不足（仅一个选择器）——抗漂移能力弱。 */
export function isAnchorWeak(anchor: NodeAnchor): boolean {
  return anchorSelectorCount(anchor) < 2;
}

/** 收集对应表记录的可变构建器。 */
export class SourceMapBuilder {
  private readonly entries: SourceMapEntry[] = [];

  /** 记录一条「语义节点 → 物理节点」的映射。 */
  add(entry: SourceMapEntry): void {
    this.entries.push(entry);
  }

  /** 已记录的条数。 */
  get size(): number {
    return this.entries.length;
  }

  /** 按稳定顺序构建对应哈希表。 */
  build(scheme: SourceMapScheme = SOURCE_MAP_SCHEME): SourceMap {
    // 排序：让同一份文档每次产出完全一致的对应表，输出可 diff、可快照。
    const sorted = [...this.entries].sort((left, right) => compareStrings(left.id, right.id));

    const bySemanticId: Record<SemanticId, string[]> = {};
    const byPointer: Record<string, SemanticId> = {};
    const byFingerprint: Record<string, SemanticId[]> = {};

    for (const entry of sorted) {
      // 正向：id -> 全部指针（段内 run 拆分时可能 1:N）。
      const pointers = bySemanticId[entry.id] ?? [];
      for (const pointer of entry.pointers) {
        if (!pointers.includes(pointer)) pointers.push(pointer);
      }
      pointers.sort(compareStrings);
      bySemanticId[entry.id] = pointers;

      // 反向：指针 -> id。首个登记者为权威，重复登记说明两个语义节点抢同一物理节点，
      // 这属于双 IR 的结构性缺陷，由 verifier 的完整性检查负责发现。
      for (const pointer of pointers) {
        if (byPointer[pointer] === undefined) {
          byPointer[pointer] = entry.id;
        }
      }

      // 内容寻址：指纹 -> id 列表（同一段文字可能出现多次，因此是 1:N）。
      const ids = byFingerprint[entry.anchor.digest] ?? [];
      if (!ids.includes(entry.id)) ids.push(entry.id);
      ids.sort(compareStrings);
      byFingerprint[entry.anchor.digest] = ids;
    }

    return { scheme, entries: sorted, bySemanticId, byPointer, byFingerprint };
  }
}

/** 字符串比较器：显式定义是为了得到与 locale 无关的稳定顺序。 */
export function compareStrings(left: string, right: string): number {
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

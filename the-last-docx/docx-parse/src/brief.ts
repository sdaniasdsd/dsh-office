/**
 * Agent 判断依据的渲染层 —— 「类似提示词设计」那部分。
 *
 * 为什么需要它，以及它为什么不算「过度设计」：
 *   双 IR 是给机器看的结构，账本是给机器查的记录，判定是给机器用的开关。
 *   但真正做决定的 Agent 需要的是【一份能读的简报】：现在是什么档位、
 *   文档里有哪些东西可以点名、用户之前说过什么、上一轮到底改了什么。
 *   把这三样拼成一段稳定文案，Agent 每次拿到的形态都一致，
 *   它的行为才可预测、可回归测试。
 *
 * 刻意的约束：
 *   1. 纯函数，不读时间、不读存储。同样的输入永远得到逐字节相同的文案。
 *   2. 不改写用户原话。原话直接进简报——改写会丢信息，而 Agent 正需要原始措辞。
 *   3. 有硬上限。文档可能有几千个节点，简报必须截断，否则塞不进上下文。
 */
import type { DocxDualIR, SemanticBlock, SemanticId } from './contract';
import type { GovernanceDecision } from './governance/contract';
import { governanceDirective } from './governance/gate';
import type { RevisionEntry, RevisionIntent } from './revision/contract';

/** 简报渲染选项。 */
export interface AgentBriefOptions {
  /** 最多列出多少个可点名目标。默认 40。 */
  maxTargets?: number;
  /** 最多列出多少条修订历史。默认 10（只列最近的）。 */
  maxHistory?: number;
  /** 单个文本片段的最大字符数。默认 60。 */
  maxSnippetLength?: number;
}

const DEFAULT_MAX_TARGETS = 40;
const DEFAULT_MAX_HISTORY = 10;
const DEFAULT_MAX_SNIPPET = 60;

/** 截断文本。 */
function snippet(text: string, limit: number): string {
  const collapsed = text.replace(/[\s\u00a0]+/g, ' ').trim();
  if (collapsed.length <= limit) return collapsed;
  return `${collapsed.slice(0, limit)}…`;
}

/**
 * 一个「可点名目标」。
 *
 * 这是简报里最关键的部分：Agent 要改东西，必须能【指名道姓】。
 * 因此每个目标都给三样东西：
 *   - id       : 写进指令里的主语，唯一且稳定
 *   - label    : 给人看的定位（H2 / 段落 / 表格单元格）
 *   - quote    : 内容摘要，让 Agent 确认自己找的是不是那一个
 */
interface Target {
  id: SemanticId;
  label: string;
  quote: string;
  pointer: string;
}

/** 收集全部可点名目标（含表格单元格内的段落）。 */
function collectTargets(content: DocxDualIR, limit: number): { targets: Target[]; total: number } {
  const all: Target[] = [];

  const addBlock = (block: SemanticBlock): void => {
    if (block.kind === 'heading') {
      all.push({
        id: block.id,
        label: `H${block.level}`,
        quote: block.text,
        pointer: block.anchor.structuralPath,
      });
    } else if (block.kind === 'paragraph') {
      all.push({
        id: block.id,
        label: '段落',
        quote: block.text,
        pointer: block.anchor.structuralPath,
      });
    } else {
      all.push({
        id: block.id,
        label: `表格 ${block.rows.length}x${block.gridColumns}`,
        quote: block.rows
          .map((row) => row.cells.map((cell) => cell.text).join(' | '))
          .join(' / '),
        pointer: block.anchor.structuralPath,
      });
      // 单元格段落单独登记：否则「改表格里某一格」无从指认。
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.paragraphs) {
            all.push({
              id: paragraph.id,
              label: '单元格段落',
              quote: paragraph.text,
              pointer: paragraph.anchor.structuralPath,
            });
          }
        }
      }
    }
  };

  for (const block of content.semantic.blocks) addBlock(block);

  return { targets: all.slice(0, limit), total: all.length };
}

/** 渲染目标清单。 */
function renderTargets(content: DocxDualIR, options: AgentBriefOptions): string[] {
  const limit = options.maxTargets ?? DEFAULT_MAX_TARGETS;
  const snippetLimit = options.maxSnippetLength ?? DEFAULT_MAX_SNIPPET;
  const { targets, total } = collectTargets(content, limit);

  if (total === 0) return ['（文档为空，没有可点名的节点）'];

  const lines = targets.map(
    (target) => `- id=${target.id} [${target.label}] ${snippet(target.quote, snippetLimit)}`,
  );
  if (total > targets.length) {
    lines.push(`… 另有 ${total - targets.length} 个节点未列出（共 ${total} 个）`);
  }
  return lines;
}

/** 渲染大纲：让 Agent 一眼看出文档结构。 */
function renderOutline(content: DocxDualIR, snippetLimit: number): string[] {
  const lines: string[] = [];
  for (const block of content.semantic.blocks) {
    if (block.kind !== 'heading') continue;
    const indent = '  '.repeat(Math.max(0, block.level - 1));
    // 推断出来的级别要标出来：它可能不准，Agent 有权知道。
    const marker = block.levelSource === 'outlineLevel' ? '' : '（级别为推断）';
    lines.push(`${indent}- H${block.level} ${snippet(block.text, snippetLimit)}${marker}`);
  }
  return lines.length === 0 ? ['（本文档没有标题层级）'] : lines;
}

/** 渲染一条修订记录的改动摘要。 */
function renderChanges(entry: RevisionEntry, snippetLimit: number): string[] {
  if (entry.changes.length === 0) {
    // 三种「零差异」含义完全不同，混成一句会让 Agent 误判成败：
    //   * 只是表达了意图        —— 文档确实没动；
    //   * 首次登记，没有前版    —— 有版本但无参照物；
    //   * 有前版但内容一致      —— 用户提了修改，结果等于没改。
    if (!entry.producedVersion) return ['  （仅记录意图，文档未变）'];
    if (entry.beforeFingerprint === null) return ['  （首个版本，无前版可比）'];
    return ['  （内容与上一版完全一致）'];
  }
  return entry.changes.slice(0, 10).map((change) => {
    const detail =
      change.change === 'modified'
        ? `字段=[${change.fields.join(',')}] ${snippet(change.before ?? '', snippetLimit)} → ${snippet(change.after ?? '', snippetLimit)}`
        : snippet(change.after ?? change.before ?? '', snippetLimit);
    return `  - ${change.change} ${change.semanticId} ${detail}`;
  });
}

/** 渲染账本。 */
function renderHistory(
  entries: readonly RevisionEntry[],
  options: AgentBriefOptions,
): string[] {
  if (entries.length === 0) return ['（尚无修订记录）'];
  const limit = options.maxHistory ?? DEFAULT_MAX_HISTORY;
  const snippetLimit = options.maxSnippetLength ?? DEFAULT_MAX_SNIPPET;

  // 只列最近的若干条：简报是给当下的决策用的，不是全量审计报告。
  const recent = entries.slice(-limit);
  const lines: string[] = [];
  if (entries.length > recent.length) {
    lines.push(`（仅列出最近 ${recent.length} 条，共 ${entries.length} 条）`);
  }
  for (const entry of recent) {
    lines.push(
      `#${entry.revision} [${entry.intent.stage}/${entry.intent.kind}] 原话：「${entry.intent.text}」`,
    );
    lines.push(...renderChanges(entry, snippetLimit));
  }
  return lines;
}

/**
 * 渲染给 Agent 的完整判断依据。
 *
 * 结构固定为五段：档位 → 纪律 → 大纲 → 可点名目标 → 账本。
 * 顺序不是随意的：先讲「这次要多严」，再给「可以动哪些东西」，
 * 最后给「之前发生过什么」，与 Agent 的决策顺序一致。
 */
export function renderAgentBrief(
  content: DocxDualIR,
  decision: GovernanceDecision,
  ledger: { intents: readonly RevisionIntent[]; entries: readonly RevisionEntry[] },
  options: AgentBriefOptions = {},
): string {
  const snippetLimit = options.maxSnippetLength ?? DEFAULT_MAX_SNIPPET;
  const sections: string[] = [];

  sections.push('## 治理档位');
  sections.push(`mode = ${decision.mode}（理由：${decision.reason}）`);
  if (decision.matchedSignals.length > 0) {
    sections.push(`命中信号：${decision.matchedSignals.join(', ')}`);
  }
  sections.push('');
  sections.push(governanceDirective(decision.mode));

  sections.push('');
  sections.push('## 文档大纲');
  sections.push(...renderOutline(content, snippetLimit));

  sections.push('');
  sections.push('## 可点名目标');
  sections.push(`视图方案：${content.scheme}｜主部件：${content.physical.mainPart}`);
  sections.push(...renderTargets(content, options));

  sections.push('');
  sections.push('## 修订账本');
  sections.push(...renderHistory(ledger.entries, options));

  // 尾注收尾：提醒 Agent 定位纪律，避免它退回「按第几段」的模糊指代。
  sections.push('');
  sections.push(
    `> 定位规则：改动一律以 id= 开头的语义节点标识指代；不要用「第 N 段」或位置序号。`,
  );

  return sections.join('\n');
}

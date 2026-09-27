/**
 * 双 IR 差异计算 —— 「新旧修改」的核心。
 *
 * 为什么这里能算得准，而通用的文本 diff 做不到？
 * 因为语义节点带有【跨版本稳定】的 id：id 由 paraId / 结构路径决定，不含内容。
 * 于是：
 *
 *   同一 id 出现在两版，指纹相同  → 没动过
 *   同一 id 出现在两版，指纹不同  → 同一个节点被改了（而不是「删掉旧的、加了新的」）
 *   id 只在旧版                  → 真的被删了
 *   id 只在新版                  → 真的被加了
 *
 * 通用文本 diff 只能看到字符层面的增删，无法区分「这句被改写了」与
 * 「这句没了、旁边多了另一句」——而这恰恰是判断用户意图的关键区别。
 *
 * 本文件是纯函数：不读时间、不读存储、无副作用。
 */
import type { ChangeField, RevisionChange } from './contract';
import type {
  DocxDualIR,
  SemanticAnnotation,
  SemanticBlock,
  SemanticParagraph,
  SemanticId,
  SemanticTable,
} from '../contract';
import { makePointer, sha256Hex, normalizeText } from '../sourcemap';

/** 摘要保留的最大字符数，避免账本体积随文档增大而失控。 */
const MAX_SUMMARY_LENGTH = 80;

/** 节点在差异计算中的可比面。 */
interface NodeFacet {
  id: SemanticId;
  kind: string;
  /**
   * 比对键。
   *
   * 段落/标题/批注用内容指纹；表格用【结构签名】而非内容——
   * 表格的内容由单元格承载，若在表格层再比一次内容，
   * 任何一格改动都会被上报两遍（表格一次、单元格段落一次）。
   * 表格层只负责回答「行列结构有没有变」。
   */
  digest: string;
  /** 样式引用；批注等非段落节点为 null。 */
  styleId: string | null;
  /** 标题级别；非标题为 null。 */
  level: number | null;
  /** 可读摘要，用于回填 before/after。 */
  summary: string;
  /** 完整物理指针（`<part>!<结构路径>`），写入端直接可用。 */
  pointer: string;
  /** 在文档中的出现次序（0 基），仅用于把结果排成可读顺序。 */
  order: number;
}

/** 表格的结构签名：只有行列形状参与比对。 */
function tableShapeDigest(block: SemanticTable): string {
  const rows = block.rows.map((row) => row.cells.length).join(',');
  return sha256Hex(`table\u0000${block.gridColumns}\u0000${rows}`);
}

/** 截断文本，供摘要使用。 */
function truncate(value: string): string {
  const normalized = normalizeText(value);
  if (normalized.length <= MAX_SUMMARY_LENGTH) return normalized;
  return `${normalized.slice(0, MAX_SUMMARY_LENGTH)}…`;
}

/** 由一个语义块构造可比面。 */
function blockFacet(block: SemanticBlock, order: number): NodeFacet {
  const facet: NodeFacet = {
    id: block.id,
    kind: block.kind,
    digest: block.anchor.digest,
    styleId: null,
    level: null,
    summary: '',
    pointer: makePointer(block.anchor.part, block.anchor.structuralPath),
    order,
  };

  if (block.kind === 'heading') {
    facet.styleId = block.styleId;
    facet.level = block.level;
    facet.summary = `H${block.level} ${truncate(block.text)}`;
  } else if (block.kind === 'paragraph') {
    facet.styleId = block.styleId;
    facet.summary = truncate(block.text);
  } else {
    // 表格层只比对结构，内容交给单元格段落。
    facet.digest = tableShapeDigest(block);
    facet.summary = `table ${block.rows.length}x${block.gridColumns}`;
  }
  return facet;
}

/** 由一个语义段落构造可比面。 */
function paragraphFacet(paragraph: SemanticParagraph, order: number): NodeFacet {
  return {
    id: paragraph.id,
    kind: paragraph.kind,
    digest: paragraph.anchor.digest,
    styleId: paragraph.styleId,
    level: null,
    summary: truncate(paragraph.text),
    pointer: makePointer(paragraph.anchor.part, paragraph.anchor.structuralPath),
    order,
  };
}

/** 由一条注释构造可比面。 */
function annotationFacet(annotation: SemanticAnnotation, order: number): NodeFacet {
  return {
    id: annotation.id,
    kind: annotation.kind,
    digest: annotation.anchor.digest,
    styleId: null,
    level: null,
    summary: truncate(annotation.text),
    pointer: makePointer(annotation.anchor.part, annotation.anchor.structuralPath),
    order,
  };
}

/**
 * 收集一份双 IR 里全部可比较节点。
 *
 * 遍历顺序即文档顺序，因此每个节点拿到的 `order` 就是它的位置序号。
 * 表格单元格内的段落【也要】收集：否则「改表格里某一格」在差异里完全不可见，
 * 而这是最常见的办公改动之一。
 */
export function collectFacets(content: DocxDualIR): Map<SemanticId, NodeFacet> {
  const facets = new Map<SemanticId, NodeFacet>();
  let order = 0;

  for (const block of content.semantic.blocks) {
    facets.set(block.id, blockFacet(block, order));
    order += 1;
    if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.paragraphs) {
            facets.set(paragraph.id, paragraphFacet(paragraph, order));
            order += 1;
          }
        }
      }
    }
  }
  for (const annotation of content.semantic.annotations) {
    facets.set(annotation.id, annotationFacet(annotation, order));
    order += 1;
  }
  return facets;
}

/** 比较两个可比面，列出变化的维度。 */
function changedFields(before: NodeFacet, after: NodeFacet): ChangeField[] {
  const fields: ChangeField[] = [];
  if (before.digest !== after.digest) {
    // 表格的比对键是结构签名，因此同样的「键变了」在表格上含义是结构变化。
    fields.push(before.kind === 'table' ? 'structure' : 'text');
  }
  if (before.styleId !== after.styleId) fields.push('style');
  if (before.level !== after.level) fields.push('level');
  return fields;
}

/** 差异计算选项。 */
export interface DiffOptions {
  /** 是否把「没变」的节点也算进结果。默认 false——账本只关心变化。 */
  includeUnchanged?: boolean;
}

/** 内部形态：带上排序键，便于最终整理成文档顺序。 */
interface OrderedChange {
  order: number;
  change: RevisionChange;
}

/**
 * 计算两份双 IR 之间的节点级差异。
 *
 * 结果按【文档顺序】排列：同一对输入永远得到完全一致的输出，
 * 但顺序对人是有意义的——第一行就是文档里最先被改动的地方。
 */
export function diffDualIR(
  before: DocxDualIR,
  after: DocxDualIR,
  options: DiffOptions = {},
): RevisionChange[] {
  const includeUnchanged = options.includeUnchanged === true;
  const beforeFacets = collectFacets(before);
  const afterFacets = collectFacets(after);

  // 并集：只在一侧出现的节点同样是重要信息（新增/删除）。
  const allIds = new Set<string>([...beforeFacets.keys(), ...afterFacets.keys()]);
  const ordered: OrderedChange[] = [];

  for (const id of allIds) {
    const previous = beforeFacets.get(id);
    const current = afterFacets.get(id);
    const kind = current?.kind ?? previous?.kind ?? 'unknown';
    // 排序键优先取新版位置（新增节点只有新版位置，删除节点只有旧版位置）。
    const order = current?.order ?? previous?.order ?? 0;

    if (previous === undefined && current !== undefined) {
      ordered.push({
        order,
        change: {
          semanticId: id,
          kind,
          change: 'added',
          fields: [],
          before: null,
          after: current.summary,
          pointers: [current.pointer],
        },
      });
      continue;
    }
    if (previous !== undefined && current === undefined) {
      ordered.push({
        order,
        change: {
          semanticId: id,
          kind,
          change: 'removed',
          fields: [],
          before: previous.summary,
          after: null,
          pointers: [previous.pointer],
        },
      });
      continue;
    }
    // 两侧都存在：比较变化维度。
    if (previous === undefined || current === undefined) continue;
    const fields = changedFields(previous, current);
    if (fields.length === 0 && !includeUnchanged) continue;
    ordered.push({
      order,
      change: {
        semanticId: id,
        kind,
        change: fields.length === 0 ? 'unchanged' : 'modified',
        fields,
        before: previous.summary,
        after: current.summary,
        pointers: [current.pointer],
      },
    });
  }

  // 同序（理论上不会发生）时退化为按 id 排，保证输出严格确定。
  return ordered
    .sort((left, right) =>
      left.order === right.order
        ? left.change.semanticId < right.change.semanticId
          ? -1
          : 1
        : left.order - right.order,
    )
    .map((entry) => entry.change);
}

/**
 * 语义视图指纹 —— 用「身份 + 可比面」算出的摘要。
 *
 * 为什么不用 IR 的完整 JSON：那样会对排序、可选字段等无关差异敏感。
 * 指纹只覆盖真正决定「文档语义状态」的那些值，因此它是判断
 * 「这次改动到底有没有改变文档」的可靠依据。
 */
export function irFingerprint(content: DocxDualIR): string {
  const facets = collectFacets(content);
  const rows = [...facets.values()]
    .map((facet) => `${facet.id}\u0000${facet.digest}\u0000${facet.styleId ?? ''}\u0000${facet.level ?? ''}`)
    .sort();
  // 视图方案一并纳入：锚点方案变了，指纹自然应该变。
  return sha256Hex(`${content.scheme}\u0001${rows.join('\u0002')}`);
}

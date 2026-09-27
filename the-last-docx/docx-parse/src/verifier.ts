/**
 * 局部验证器 —— 在【本模块能力范围内】对双 IR 与安全策略做检查。
 *
 * 职责边界（spec 第八节「Profile 集成要求」）：
 *   - 这里只做「结构 + 双 IR 一致性 + 策略」层面的验证，属于 docx-parse 的能力边界内；
 *   - 视觉/渲染类验证由 office-preview、共享验证套件由 office-test-kit 负责；
 *   - 报告结构由本模块拥有，底层库无法改写。
 *
 * 为什么这些检查特别重要？
 *   回写场景没有「权威裁判」——没有任何开源实现能声称自己与 Word 渲染完全一致。
 *   因此我们能拿到的最强证据，是两条【可判定的不变量】：
 *     1. 对应表自洽（语义节点与物理节点严丝合缝，没有孤儿、没有抢占）；
 *     2. 身份确定性（同一份输入重算 id 必须逐字节相同）。
 *   它们不能替代「解析对不对」，但能排除「双 IR 内部对不上号」这一整类致命缺陷。
 *
 * 三态语义（贯穿全部检查）：
 *   pass —— 策略声明了该要求，且产物满足；
 *   fail —— 策略声明了该要求，但产物不满足；
 *   skip —— 策略【未声明】该要求，因此无法判断。
 */
import type { SafetyPolicy } from 'office-safety';

import type {
  DocxDualIR,
  JsonObject,
  ModuleConfig,
  VerificationCheck,
  VerificationReport,
  VerificationSummary,
} from './contract';
import type { ParseResult } from './domain/docx-parse';
import { computeSemanticId } from './sourcemap';

/** 检查项 id 常量表：稳定 id 是回归测试能做精确断言的前提。 */
export const VERIFICATION_CHECK_IDS = {
  STRUCTURE: 'structure.package',
  SEMANTIC: 'content.semanticView',
  SOURCEMAP: 'sourcemap.integrity',
  DETERMINISM: 'identity.deterministic',
  SERIALIZABLE: 'output.serializable',
  ENCRYPTION: 'policy.encryption',
  BLOCKS: 'policy.blockBudget',
  TABLE_CELLS: 'policy.tableCellBudget',
} as const;

/** 构造一条检查项；details 为空时省略该字段，保持输出紧凑。 */
function check(
  id: string,
  status: VerificationCheck['status'],
  severity: VerificationCheck['severity'],
  message: string,
  details?: JsonObject,
): VerificationCheck {
  const entry: VerificationCheck = { id, status, severity, message };
  if (details !== undefined && Object.keys(details).length > 0) {
    entry.details = details;
  }
  return entry;
}

/** 安全策略布尔项的通用检查（未声明或未解析即 skip）。 */
function policyBooleanCheck(
  id: string,
  allowed: boolean | undefined,
  present: boolean,
  parsed: boolean,
  presentMessage: string,
  absentMessage: string,
): VerificationCheck {
  if (!parsed) {
    return check(id, 'skip', 'info', 'Indicator was not parsed in this call (feature flag disabled).');
  }
  if (allowed === undefined) {
    return check(id, 'skip', 'info', 'Policy does not state a requirement for this indicator.');
  }
  if (!present) {
    return check(id, 'pass', 'info', absentMessage);
  }
  if (allowed) {
    return check(id, 'pass', 'info', presentMessage);
  }
  return check(id, 'fail', 'error', presentMessage);
}

/** 结构完整性检查（唯一的「内部」前置检查，不受策略开关影响）。 */
function structuralCheck(result: ParseResult): VerificationCheck {
  const problems: string[] = [];
  if (result.container !== 'zip') {
    problems.push(`container is ${result.container}`);
  }
  if (result.documentKind !== 'wordprocessingml') {
    problems.push(`documentKind is ${result.documentKind}`);
  }
  if (result.partCount === 0) {
    problems.push('package has no parts');
  }

  if (problems.length > 0) {
    return check(VERIFICATION_CHECK_IDS.STRUCTURE, 'fail', 'error', 'Package structure is not valid.', {
      problems,
    });
  }
  return check(
    VERIFICATION_CHECK_IDS.STRUCTURE,
    'pass',
    'info',
    'Package is a well-formed WordprocessingML container.',
    { partCount: result.partCount, xmlBackend: result.xmlBackend },
  );
}

/**
 * 语义视图自检：id 唯一、标题级别合法、表格行列结构不空。
 *
 * 这是「双 IR 的第一份视图是否自洽」的检查，与对应表检查互为补充。
 */
function semanticViewCheck(content: DocxDualIR): VerificationCheck {
  const problems: string[] = [];
  const seen = new Set<string>();

  const visitParagraphId = (id: string): void => {
    if (seen.has(id)) problems.push(`duplicate semantic id ${id}`);
    seen.add(id);
  };

  for (const block of content.semantic.blocks) {
    if (seen.has(block.id)) problems.push(`duplicate semantic id ${block.id}`);
    seen.add(block.id);

    if (block.kind === 'heading') {
      if (!Number.isInteger(block.level) || block.level < 1 || block.level > 9) {
        problems.push(`heading ${block.id} has out-of-range level ${block.level}`);
      }
    } else if (block.kind === 'table') {
      if (block.rows.length === 0) {
        problems.push(`table ${block.id} has no rows`);
      }
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.paragraphs) {
            visitParagraphId(paragraph.id);
          }
        }
      }
    }
  }

  for (const annotation of content.semantic.annotations) {
    if (seen.has(annotation.id)) problems.push(`duplicate semantic id ${annotation.id}`);
    seen.add(annotation.id);
  }

  if (problems.length > 0) {
    return check(
      VERIFICATION_CHECK_IDS.SEMANTIC,
      'fail',
      'error',
      'Semantic view is not internally consistent.',
      { problems: problems.slice(0, 20) },
    );
  }
  return check(VERIFICATION_CHECK_IDS.SEMANTIC, 'pass', 'info', 'Semantic view is internally consistent.', {
    blockCount: content.semantic.blocks.length,
    annotationCount: content.semantic.annotations.length,
  });
}

/** 收集双 IR 中出现的全部语义 id（含表格单元格内的段落）。 */
function collectSemanticIds(content: DocxDualIR): string[] {
  const ids: string[] = [];
  for (const block of content.semantic.blocks) {
    ids.push(block.id);
    if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.paragraphs) {
            ids.push(paragraph.id);
          }
        }
      }
    }
  }
  for (const annotation of content.semantic.annotations) {
    ids.push(annotation.id);
  }
  return ids;
}

/**
 * 对应表完整性检查 —— 双 IR 的核心不变量。
 *
 * 检查四件事：
 *   1. 每个语义节点都能在 bySemanticId 里找到；
 *   2. 每条 entry 的 pointer 确实登记在 bySemanticId[entry.id] 中；
 *   3. 反向索引 byPointer 与 entry 一致；
 *   4. 没有任何物理指针被两个语义节点同时认领（1:1 的物理落点不能被瓜分）。
 */
function sourceMapCheck(content: DocxDualIR, config: ModuleConfig): VerificationCheck {
  if (!config.featureFlags.resolveAnchors) {
    return check(
      VERIFICATION_CHECK_IDS.SOURCEMAP,
      'skip',
      'info',
      'Anchors were not resolved in this call (resolveAnchors disabled).',
    );
  }

  const problems: string[] = [];
  const claimByPointer = new Map<string, string>();

  for (const entry of content.sourceMap.entries) {
    const pointers = content.sourceMap.bySemanticId[entry.id];
    if (pointers === undefined) {
      problems.push(`entry ${entry.id} is missing from bySemanticId`);
    } else if (!pointers.includes(entry.pointer)) {
      problems.push(`entry ${entry.id} primary pointer ${entry.pointer} not listed in bySemanticId`);
    }
    for (const pointer of entry.pointers) {
      const owner = claimByPointer.get(pointer);
      if (owner !== undefined && owner !== entry.id) {
        problems.push(`pointer ${pointer} is claimed by both ${owner} and ${entry.id}`);
      }
      claimByPointer.set(pointer, entry.id);
      if (content.sourceMap.byPointer[pointer] !== entry.id) {
        problems.push(`byPointer[${pointer}] does not resolve back to ${entry.id}`);
      }
    }
  }

  const missing = collectSemanticIds(content).filter(
    (id) => content.sourceMap.bySemanticId[id] === undefined,
  );
  if (missing.length > 0) {
    problems.push(`${missing.length} semantic node(s) have no source-map entry`);
  }

  if (problems.length > 0) {
    return check(
      VERIFICATION_CHECK_IDS.SOURCEMAP,
      'fail',
      'error',
      'Source map between the two views is not consistent.',
      { problems: problems.slice(0, 20) },
    );
  }
  return check(
    VERIFICATION_CHECK_IDS.SOURCEMAP,
    'pass',
    'info',
    'Every semantic node resolves to exactly one physical anchor.',
    { entries: content.sourceMap.entries.length },
  );
}

/**
 * 身份确定性检查 —— 在验证期把 id 重算一遍。
 *
 * 这是「没有权威裁判」时能拿到的最强证据之一：如果 id 算法有任何不确定性
 * （例如依赖了遍历顺序、时间戳、随机数），重算就会与既有 id 不一致，
 * 从而在这里被立刻抓住。
 */
function identityDeterminismCheck(content: DocxDualIR, config: ModuleConfig): VerificationCheck {
  if (!config.featureFlags.resolveAnchors) {
    return check(
      VERIFICATION_CHECK_IDS.DETERMINISM,
      'skip',
      'info',
      'Anchors were not resolved in this call (resolveAnchors disabled).',
    );
  }

  const mismatches: string[] = [];
  for (const entry of content.sourceMap.entries) {
    const recomputed = computeSemanticId(entry.anchor);
    if (recomputed !== entry.id) {
      mismatches.push(`${entry.id} -> ${recomputed}`);
    }
  }

  if (mismatches.length > 0) {
    return check(
      VERIFICATION_CHECK_IDS.DETERMINISM,
      'fail',
      'error',
      'Recomputing node identities did not reproduce the stored ids.',
      { mismatches: mismatches.slice(0, 20) },
    );
  }
  return check(
    VERIFICATION_CHECK_IDS.DETERMINISM,
    'pass',
    'info',
    'Recomputing node identities reproduces every stored id.',
    { entries: content.sourceMap.entries.length },
  );
}

/** 可序列化性检查（运行时守住「输出可 JSON 序列化」这条不变量）。 */
function serializableCheck(
  result: ParseResult,
  content: DocxDualIR,
): VerificationCheck {
  try {
    JSON.stringify({ result, content });
    return check(
      VERIFICATION_CHECK_IDS.SERIALIZABLE,
      'pass',
      'info',
      'Verification inputs serialize to JSON.',
    );
  } catch (error) {
    return check(
      VERIFICATION_CHECK_IDS.SERIALIZABLE,
      'fail',
      'error',
      'Verification inputs are not JSON-serializable.',
      { reason: error instanceof Error ? error.message : String(error) },
    );
  }
}

/** 统计表格单元格数量。 */
function countTableCells(content: DocxDualIR): number {
  let count = 0;
  for (const block of content.semantic.blocks) {
    if (block.kind !== 'table') continue;
    for (const row of block.rows) {
      count += row.cells.length;
    }
  }
  return count;
}

/**
 * 执行局部验证，产出报告。
 *
 * 注意执行顺序：结构检查在前（前置条件），双 IR 一致性居中，策略检查与序列化自检收尾。
 * 这样当结构不成立时，阅读者能立刻看出「根因在格式，而不在策略或对应表」。
 */
export function runVerification(
  result: ParseResult,
  content: DocxDualIR,
  policy: SafetyPolicy,
  config: ModuleConfig,
): VerificationReport {
  const checks: VerificationCheck[] = [];

  // --- 1. 前置：结构完整性 ---------------------------------------------- //
  checks.push(structuralCheck(result));

  // --- 2. 双 IR 自洽性 --------------------------------------------------- //
  checks.push(semanticViewCheck(content));
  checks.push(sourceMapCheck(content, config));
  checks.push(identityDeterminismCheck(content, config));

  // --- 3. 策略驱动的检查 ------------------------------------------------ //
  checks.push(
    policyBooleanCheck(
      VERIFICATION_CHECK_IDS.ENCRYPTION,
      policy.allowEncrypted,
      result.encrypted === true,
      // 加密标记来自 ZIP 结构本身，不受能力开关影响。
      true,
      'Artifact is encrypted.',
      'Artifact is not encrypted.',
    ),
  );

  // 正文块数量预算：仅当策略给了上限时才有意义。
  if (policy.maxBlocks === undefined) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.BLOCKS, 'skip', 'info', 'Policy does not set a block budget.'),
    );
  } else {
    const count = content.semantic.blocks.length;
    const withinBudget = count <= policy.maxBlocks;
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.BLOCKS,
        withinBudget ? 'pass' : 'fail',
        withinBudget ? 'info' : 'error',
        `Document has ${count} block(s); policy budget is ${policy.maxBlocks}.`,
        { count, budget: policy.maxBlocks },
      ),
    );
  }

  // 表格单元格数量预算。
  if (policy.maxTableCells === undefined) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.TABLE_CELLS, 'skip', 'info', 'Policy does not set a table-cell budget.'),
    );
  } else {
    const count = countTableCells(content);
    const withinBudget = count <= policy.maxTableCells;
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.TABLE_CELLS,
        withinBudget ? 'pass' : 'fail',
        withinBudget ? 'info' : 'error',
        `Document has ${count} table cell(s); policy budget is ${policy.maxTableCells}.`,
        { count, budget: policy.maxTableCells },
      ),
    );
  }

  // --- 4. 收尾：序列化自检 ---------------------------------------------- //
  checks.push(serializableCheck(result, content));

  // 注意：serializable 检查本身也会改变统计，因此必须在加入它之后再汇总。
  const summary = summarize(checks);
  return {
    policyId: policy.id,
    ok: summary.failed === 0,
    partial: summary.skipped > 0,
    checks,
    summary,
  };
}

/** 统计各状态数量。 */
function summarize(checks: VerificationCheck[]): VerificationSummary {
  let passed = 0;
  let failed = 0;
  let skipped = 0;
  for (const entry of checks) {
    if (entry.status === 'pass') passed += 1;
    else if (entry.status === 'fail') failed += 1;
    else skipped += 1;
  }
  return { total: checks.length, passed, failed, skipped };
}

/**
 * 判断报告是否包含「硬失败」（error 级 fail）。
 *
 * 用途：`inspect` 的快速失败策略使用它——info 级失败不应阻止 inspect 返回结果，
 * 但 error 级失败（如结构不成立、对应表对不上号）应当阻止。
 */
export function hasBlockingFailure(report: VerificationReport): boolean {
  return report.checks.some((entry) => entry.status === 'fail' && entry.severity === 'error');
}

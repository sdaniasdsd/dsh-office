/**
 * verifier —— 本模块的【局部验证】。
 *
 * 它回答的不是「文档对不对」，而是三个更小、更能给出确定答案的问题：
 *
 *   1. **我们的产物自洽吗**（结构不变式、可序列化）—— 与文档无关，纯粹查自己；
 *   2. **策略声明的要求满足吗**（宏、外部链接、加密、预算）—— 只看调用方给了什么；
 *   3. **有没有低于置信度下限的结论** —— 不是失败，但必须让上层知道。
 *
 * 三态语义是这份报告的关键：`pass` / `fail` / `skip`。`skip` 表示**策略没有声明
 * 该要求**，而不是「通过了」。把没表态当成通过，是安全策略里最常见的错误。
 */
import type { SafetyPolicy } from 'office-safety';
import { posix } from 'node:path';

import type {
  DocxComplexParseIR,
  JsonObject,
  VerificationCheck,
  VerificationReport,
  VerificationSummary,
  SourceCoordinate,
} from './contract';
import { countTableCells } from './mapper';

/**
 * 检查项 id 常量表。
 *
 * 用常量而不是散落的字符串字面量：这些 id 会出现在快照测试与下游的规则里，
 * 一旦改名就应当是一次编译期可见的改动。
 */
export const VERIFICATION_CHECK_IDS = {
  STRUCTURE_PACKAGE: 'structure.package',
  CONTENT_REFERENCES: 'content.references',
  CONTENT_TABLES: 'content.tables',
  CONTENT_COORDINATES: 'content.coordinates',
  CONFIDENCE_FLOOR: 'confidence.floor',
  OUTPUT_SERIALIZABLE: 'output.serializable',
  POLICY_MACROS: 'policy.macros',
  POLICY_EXTERNAL_LINKS: 'policy.externalLinks',
  POLICY_ENCRYPTION: 'policy.encryption',
  POLICY_BLOCK_BUDGET: 'policy.blockBudget',
  POLICY_TABLE_CELL_BUDGET: 'policy.tableCellBudget',
} as const;

export type VerificationCheckId =
  (typeof VERIFICATION_CHECK_IDS)[keyof typeof VERIFICATION_CHECK_IDS];

/** 验证所需的一切。刻意不接收引擎对象——验证器只看得懂 IR。 */
export interface VerificationInput {
  ir: DocxComplexParseIR;
  /** 未提供时，全部策略相关检查记为 `skip`。 */
  policy?: SafetyPolicy;
  confidenceFloor: number;
}

function check(
  id: VerificationCheckId,
  passed: boolean,
  message: string,
  severity: VerificationCheck['severity'],
  details?: JsonObject,
): VerificationCheck {
  const base: VerificationCheck = {
    id,
    status: passed ? 'pass' : 'fail',
    severity,
    message,
  };
  return passed || details === undefined ? base : { ...base, details };
}

/** 策略未声明该要求 → `skip`，而不是「当作通过」。 */
function skipped(id: VerificationCheckId, requirement: string): VerificationCheck {
  return {
    id,
    status: 'skip',
    severity: 'info',
    message: `The policy does not declare ${requirement}; nothing to verify.`,
  };
}

/* -------------------------------------------------------------------------- */
/* 自洽性检查                                                                  */
/* -------------------------------------------------------------------------- */

function checkPackageStructure(ir: DocxComplexParseIR): VerificationCheck {
  const { content, parts } = ir;
  if (content.mainPart.length === 0) {
    return check(
      VERIFICATION_CHECK_IDS.STRUCTURE_PACKAGE,
      false,
      'The IR does not name a main document part.',
      'error',
    );
  }
  const present = parts.some((part) => part.name === content.mainPart);
  return check(
    VERIFICATION_CHECK_IDS.STRUCTURE_PACKAGE,
    present,
    present
      ? `Main part ${content.mainPart} is listed in the package.`
      : `Main part ${content.mainPart} is not listed in the package.`,
    'error',
    { mainPart: content.mainPart, partCount: parts.length },
  );
}

function checkTables(ir: DocxComplexParseIR): VerificationCheck {
  for (const block of ir.content.blocks) {
    const table = block.table;
    if (!table) continue;
    if (table.grid.length !== table.rowCount) {
      return check(
        VERIFICATION_CHECK_IDS.CONTENT_TABLES,
        false,
        `Table ${table.id} has ${table.grid.length} row(s) in its grid but rowCount=${table.rowCount}.`,
        'error',
        { tableId: table.id, gridRows: table.grid.length, rowCount: table.rowCount },
      );
    }
    for (const line of table.grid) {
      if (line.length !== table.columnCount) {
        return check(
          VERIFICATION_CHECK_IDS.CONTENT_TABLES,
          false,
          `Table ${table.id} has a row with ${line.length} grid cell(s) but columnCount=${table.columnCount}.`,
          'error',
          { tableId: table.id, rowLength: line.length, columnCount: table.columnCount },
        );
      }
    }
  }
  return check(
    VERIFICATION_CHECK_IDS.CONTENT_TABLES,
    true,
    'Every logical grid is a rowCount × columnCount rectangle.',
    'error',
  );
}

function checkCoordinates(ir: DocxComplexParseIR): VerificationCheck {
  for (const block of ir.content.blocks) {
    const { part, structuralPath } = block.coordinate;
    if (part.length === 0 || structuralPath.length === 0) {
      return check(
        VERIFICATION_CHECK_IDS.CONTENT_COORDINATES,
        false,
        `Block ${block.id} has no structural coordinate.`,
        'error',
        { blockId: block.id },
      );
    }
  }
  return check(
    VERIFICATION_CHECK_IDS.CONTENT_COORDINATES,
    true,
    'Every block carries a structural coordinate.',
    'error',
  );
}

/** Validate joins between package facts, logical nodes and physical fragments. */
function checkReferences(ir: DocxComplexParseIR): VerificationCheck {
  const fail = (message: string) => check(VERIFICATION_CHECK_IDS.CONTENT_REFERENCES, false, message, 'error');
  const parts = new Set(ir.parts.map((part) => part.name));
  const nodes = new Map(ir.content.blocks.map((block) => [block.id, block]));
  const floats = new Set(ir.content.floats.map((item) => item.id));
  const pages = new Map(ir.content.pages.map((page) => [page.index, page]));
  const nodeOrder = new Map(ir.content.blocks.map((block, index) => [block.id, index]));
  const pageNodes = new Map(ir.content.pages.map((page) => [page.index, new Set(page.nodeIds)]));
  const relationships = new Map<string, typeof ir.relationships>();
  const relationshipKey = (part: string, id: string) => JSON.stringify([part, id]);
  for (const rel of ir.relationships) {
    const key = relationshipKey(rel.sourcePart, rel.id);
    const matches = relationships.get(key) ?? [];
    matches.push(rel);
    relationships.set(key, matches);
  }
  if (nodes.size !== ir.content.blocks.length || floats.size !== ir.content.floats.length
    || [...floats].some((id) => nodes.has(id)) || pages.size !== ir.content.pages.length) {
    return fail('Node IDs and page indices must be unique.');
  }
  const validCoordinate = (coordinate: SourceCoordinate) => parts.has(coordinate.part)
    && coordinate.structuralPath.length > 0
    && (coordinate.pageIndex === null || pages.has(coordinate.pageIndex));
  for (const block of nodes.values()) {
    if (!validCoordinate(block.coordinate)) return fail(`Block ${block.id} references a missing part or page.`);
    const fragments = block.fragments;
    for (const fragment of fragments ?? []) {
      if (!validCoordinate(fragment) || fragment.part !== block.coordinate.part
        || fragment.structuralPath !== block.coordinate.structuralPath || fragment.pageIndex === null
        || !pageNodes.get(fragment.pageIndex)?.has(block.id)) {
        return fail(`Block ${block.id} has an unindexed or inconsistent layout fragment.`);
      }
    }
    if (block.coordinate.pageIndex !== null && !pageNodes.get(block.coordinate.pageIndex)?.has(block.id)) {
      return fail(`Block ${block.id} is absent from its primary page index.`);
    }
    const table = block.table;
    if (!table) continue;
    if (table.id !== block.id || !validCoordinate(table.coordinate)) return fail(`Table ${table.id} is detached from its block.`);
    for (const cell of table.grid.flat()) {
      if (cell.coordinate && !validCoordinate(cell.coordinate)) return fail(`Table ${table.id} has a dangling cell coordinate.`);
    }
    if (fragments !== undefined) {
      const occupied = [...new Set(fragments.flatMap((fragment) => fragment.pageIndex === null ? [] : [fragment.pageIndex]))].sort((a, b) => a - b);
      const expectedRange = occupied.length ? [occupied[0], occupied[occupied.length - 1]] : null;
      if (JSON.stringify(table.pageRange) !== JSON.stringify(expectedRange)
        || table.spansPages !== (occupied.length ? occupied.length > 1 : null)) {
        return fail(`Table ${table.id} has a page range inconsistent with its fragments.`);
      }
    }
  }
  for (const page of pages.values()) {
    if (new Set(page.nodeIds).size !== page.nodeIds.length) return fail(`Page ${page.index} repeats node IDs.`);
    let previous = -1;
    for (const id of page.nodeIds) {
      const node = nodes.get(id);
      if (!node) return fail(`Page ${page.index} references missing block ${id}.`);
      const order = nodeOrder.get(id)!;
      if (order < previous) return fail(`Page ${page.index} has inconsistent reading order.`);
      previous = order;
      if (node.fragments !== undefined && !node.fragments.some((f) => f.pageIndex === page.index)) {
        return fail(`Page ${page.index} lists block ${id} without a fragment on that page.`);
      }
    }
  }
  for (const item of ir.content.breaks) {
    if (!validCoordinate(item.coordinate) || (item.beforeNodeId !== null && !nodes.has(item.beforeNodeId))
      || [item.pageIndexBefore, item.pageIndexAfter].some((index) => index !== null && !pages.has(index))) {
      return fail('A page break references a missing node, part or page.');
    }
  }
  for (const [index, section] of ir.content.sections.entries()) {
    if (section.index !== index || (section.startBeforeNodeId !== null && !nodes.has(section.startBeforeNodeId))) {
      return fail('A section has an invalid index or start node.');
    }
  }
  for (const item of ir.content.floats) {
    if (!validCoordinate(item.coordinate) || item.part !== item.coordinate.part) return fail(`Drawing ${item.id} references a missing part.`);
    if (item.relationshipId === null) continue;
    const matches = relationships.get(relationshipKey(item.part, item.relationshipId)) ?? [];
    if (matches.length !== 1) return fail(`Drawing ${item.id} needs one relationship in its source part.`);
    const rel = matches[0]!;
    if (rel.targetMode === 'External') continue;
    let target: string;
    try { target = decodeURIComponent(rel.target.split('#')[0]!); }
    catch { return fail(`Drawing ${item.id} has an invalid relationship target.`); }
    target = target.startsWith('/') ? posix.normalize(target).slice(1)
      : posix.normalize(posix.join(posix.dirname(item.part), target));
    if (!parts.has(target)) return fail(`Drawing ${item.id} points to missing package part ${target}.`);
  }
  return check(VERIFICATION_CHECK_IDS.CONTENT_REFERENCES, true, 'Package, content and page references are consistent.', 'error');
}

function checkConfidenceFloor(input: VerificationInput): VerificationCheck {
  const { ir, confidenceFloor } = input;
  let below = 0;
  for (const block of ir.content.blocks) {
    if (block.confidence.score < confidenceFloor) below += 1;
    for (const line of block.table?.grid ?? []) {
      for (const cell of line) {
        if (cell.confidence.score < confidenceFloor) below += 1;
      }
    }
  }
  for (const item of [...ir.content.breaks, ...ir.content.floats]) {
    if (item.confidence.score < confidenceFloor) below += 1;
  }
  return check(
    VERIFICATION_CHECK_IDS.CONFIDENCE_FLOOR,
    below === 0,
    below === 0
      ? `Every conclusion reaches the confidence floor ${confidenceFloor}.`
      : `${below} conclusion(s) fall below the confidence floor ${confidenceFloor}.`,
    // 低置信度不是一个错误：它是一条「请复核」的提示，不该阻断流程。
    'warn',
    { below, confidenceFloor },
  );
}

function checkSerializable(ir: DocxComplexParseIR): VerificationCheck {
  try {
    JSON.stringify(ir);
    return check(
      VERIFICATION_CHECK_IDS.OUTPUT_SERIALIZABLE,
      true,
      'The IR round-trips through JSON.',
      'error',
    );
  } catch (error) {
    return check(
      VERIFICATION_CHECK_IDS.OUTPUT_SERIALIZABLE,
      false,
      'The IR could not be serialized to JSON.',
      'error',
      { reason: error instanceof Error ? error.message : String(error) },
    );
  }
}

/* -------------------------------------------------------------------------- */
/* 策略检查                                                                    */
/* -------------------------------------------------------------------------- */

function checkPolicyMacros(input: VerificationInput): VerificationCheck {
  const allow = input.policy?.allowMacros;
  if (allow === undefined) return skipped(VERIFICATION_CHECK_IDS.POLICY_MACROS, 'allowMacros');
  const macros = input.ir.indicators.macros;
  return check(
    VERIFICATION_CHECK_IDS.POLICY_MACROS,
    allow || macros.length === 0,
    allow
      ? 'The policy allows macros.'
      : macros.length === 0
        ? 'No macro indicator was found.'
        : `The policy forbids macros but ${macros.length} indicator(s) were found.`,
    'error',
    { macroCount: macros.length },
  );
}

function checkPolicyExternalLinks(input: VerificationInput): VerificationCheck {
  const allow = input.policy?.allowExternalLinks;
  if (allow === undefined)
    return skipped(VERIFICATION_CHECK_IDS.POLICY_EXTERNAL_LINKS, 'allowExternalLinks');
  const links = input.ir.indicators.externalReferences;
  return check(
    VERIFICATION_CHECK_IDS.POLICY_EXTERNAL_LINKS,
    allow || links.length === 0,
    allow
      ? 'The policy allows external links.'
      : links.length === 0
        ? 'No external reference was found.'
        : `The policy forbids external links but ${links.length} reference(s) were found.`,
    'error',
    { externalReferenceCount: links.length },
  );
}

function checkPolicyEncryption(input: VerificationInput): VerificationCheck {
  const allow = input.policy?.allowEncrypted;
  if (allow === undefined)
    return skipped(VERIFICATION_CHECK_IDS.POLICY_ENCRYPTION, 'allowEncrypted');

  const encrypted = input.ir.profile.encrypted;
  if (allow) {
    return check(
      VERIFICATION_CHECK_IDS.POLICY_ENCRYPTION,
      true,
      'The policy allows encrypted artifacts.',
      'error',
      { encrypted },
    );
  }
  // 失败方向朝「拦」而不是朝「放」：策略说不允许加密，而加密状态【无法判断】时，
  // 我们并没有资格宣布它合规——按不满足处理，把决定权交回策略作者。
  const passed = encrypted === false;
  return check(
    VERIFICATION_CHECK_IDS.POLICY_ENCRYPTION,
    passed,
    passed
      ? 'The artifact is confirmed unencrypted.'
      : encrypted === true
        ? 'The policy forbids encrypted artifacts but this one is encrypted.'
        : 'The policy forbids encrypted artifacts but the encryption state is unknown.',
    'error',
    { encrypted },
  );
}

function checkPolicyBlockBudget(input: VerificationInput): VerificationCheck {
  const max = input.policy?.maxBlocks;
  if (max === undefined) return skipped(VERIFICATION_CHECK_IDS.POLICY_BLOCK_BUDGET, 'maxBlocks');
  const count = input.ir.content.blocks.length;
  return check(
    VERIFICATION_CHECK_IDS.POLICY_BLOCK_BUDGET,
    count <= max,
    count <= max
      ? `${count} block(s) is within the policy budget ${max}.`
      : `${count} block(s) exceeds the policy budget ${max}.`,
    'error',
    { blockCount: count, maxBlocks: max },
  );
}

function checkPolicyTableCellBudget(input: VerificationInput): VerificationCheck {
  const max = input.policy?.maxTableCells;
  if (max === undefined)
    return skipped(VERIFICATION_CHECK_IDS.POLICY_TABLE_CELL_BUDGET, 'maxTableCells');
  const count = countTableCells(input.ir.content);
  return check(
    VERIFICATION_CHECK_IDS.POLICY_TABLE_CELL_BUDGET,
    count <= max,
    count <= max
      ? `${count} table cell(s) is within the policy budget ${max}.`
      : `${count} table cell(s) exceeds the policy budget ${max}.`,
    'error',
    { tableCellCount: count, maxTableCells: max },
  );
}

/* -------------------------------------------------------------------------- */
/* 汇总                                                                        */
/* -------------------------------------------------------------------------- */

function summarize(checks: readonly VerificationCheck[]): VerificationSummary {
  let passed = 0;
  let failed = 0;
  let skippedCount = 0;
  for (const item of checks) {
    if (item.status === 'pass') passed += 1;
    else if (item.status === 'fail') failed += 1;
    else skippedCount += 1;
  }
  return { total: checks.length, passed, failed, skipped: skippedCount };
}

/**
 * 执行局部验证。
 *
 * 纯函数：同样的输入必然产出同样的报告（检查顺序固定，不依赖时间或随机数）。
 */
export function verify(input: VerificationInput): VerificationReport {
  const checks: VerificationCheck[] = [
    checkPackageStructure(input.ir),
    checkTables(input.ir),
    checkCoordinates(input.ir),
    checkReferences(input.ir),
    checkSerializable(input.ir),
    checkConfidenceFloor(input),
    checkPolicyMacros(input),
    checkPolicyExternalLinks(input),
    checkPolicyEncryption(input),
    checkPolicyBlockBudget(input),
    checkPolicyTableCellBudget(input),
  ];

  const summary = summarize(checks);
  const failed = checks.filter((item) => item.status === 'fail');

  const report: VerificationReport = {
    // `ok` 只要没有失败项就为真；`severity` 决定该失败是否【阻断】。
    // 两者分开，是为了让「低置信度」这类提示不把流程拦下来。
    ok: failed.length === 0,
    // `partial` 沿用 office-core 的定义：**验证**本身不完整——有检查被跳过
    // （通常是策略没声明对应要求），而不是「文档只被建模了一部分」。
    // 后者由 PARTIALLY_PARSED 告警表达，两者不可混用。
    partial: summary.skipped > 0,
    checks,
    summary,
  };

  return input.policy === undefined ? report : { ...report, policyId: input.policy.id };
}

/**
 * 是否存在阻断性失败。
 *
 * `inspect` 用它做快速失败：`warn` 级的失败（低置信度）不阻断，
 * `error` 级的失败必须让调用方立刻知道。
 */
export function hasBlockingFailure(report: VerificationReport): boolean {
  return report.checks.some((item) => item.status === 'fail' && item.severity === 'error');
}

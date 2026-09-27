/**
 * 局部验证器 —— 在【本模块能力范围内】对解析产物做结构与质量检查。
 *
 * 职责边界（spec 第八节「Profile 集成要求」）：
 *   - 这里只做「结构完整性 + 解析质量」层面的验证，属于 docx-parse 的能力边界内；
 *   - 视觉/渲染类验证由 office-preview、共享报告链路由 office-test-kit 负责，
 *     那些不在本模块范围内；
 *   - 报告结构由本模块拥有，底层库无法改写。
 *
 * 三态语义（贯穿全部检查）：
 *   pass —— 策略声明了该要求，且产物满足；
 *   fail —— 策略声明了该要求，但产物不满足；
 *   skip —— 策略【未声明】该要求（或能力开关关闭导致无法判断）。
 *
 * 之所以要区分 skip 与 pass：策略作者必须能表达「我确认允许」与
 * 「我根本没表态」这两种截然不同的意思，否则报告会产生虚假的安全感。
 */
import type {
  JsonObject,
  ModuleConfig,
  ParsePolicy,
  VerificationCheck,
  VerificationReport,
  VerificationSummary,
} from './contract';
import { computeCounts, type ParseResult } from './domain/docx-parse';

/** 检查项 id 常量表：稳定 id 是回归测试能做精确断言的前提。 */
export const VERIFICATION_CHECK_IDS = {
  STRUCTURE: 'structure.document',
  SERIALIZABLE: 'output.serializable',
  ENCRYPTION: 'policy.encryption',
  HEADINGS: 'policy.headings',
  MIN_PARAGRAPHS: 'policy.minParagraphs',
  STYLES: 'policy.styles',
  RESOLVED_STYLES: 'policy.resolvedStyles',
  DANGLING_RELATIONSHIPS: 'policy.danglingRelationships',
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

/** 未声明要求 / 未探测指标的统一 skip。 */
function skipped(id: string, reason: string): VerificationCheck {
  return check(id, 'skip', 'info', reason);
}

/** 判断解析结果里是否出现了某个引擎问题码。 */
function hasIssue(result: ParseResult, code: string): boolean {
  return result.issues.some((issue) => issue.code === code);
}

/**
 * 结构完整性检查。
 *
 * 这是唯一的「内部」检查（不受策略开关影响）：如果文档结构本身不成立，
 * 那么任何策略结论都是不可靠的，必须无条件先确认。
 */
function structuralCheck(result: ParseResult): VerificationCheck {
  const problems: string[] = [];
  if (result.container !== 'zip') {
    problems.push(`container is ${result.container}`);
  }
  if (result.documentKind !== 'wordprocessingml') {
    problems.push(`documentKind is ${result.documentKind}`);
  }

  if (problems.length > 0) {
    return check(
      VERIFICATION_CHECK_IDS.STRUCTURE,
      'fail',
      'error',
      'Document structure is not valid.',
      { problems },
    );
  }
  const counts = computeCounts(result);
  return check(
    VERIFICATION_CHECK_IDS.STRUCTURE,
    'pass',
    'info',
    'Document is a well-formed WordprocessingML package.',
    { blocks: counts.blocks, xmlBackend: result.xmlBackend },
  );
}

/**
 * 可序列化性检查。
 *
 * 为什么把它做成「检查项」而不是只在测试里断言：spec 把
 * 「输入输出错误 warning 均可 JSON 序列化」列为完成标准，
 * 那么在运行时自检一次，就能在真实调用链路上持续守住这条不变量，
 * 而不是依赖某次测试恰好跑到。
 */
function serializableCheck(result: ParseResult, checks: VerificationCheck[]): VerificationCheck {
  try {
    JSON.stringify({ result, checks });
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

/**
 * 执行局部验证，产出报告。
 *
 * 注意执行顺序：结构检查在前（前置条件），策略检查在后；
 * 这样当结构不成立时，阅读者能立刻看出「根因在格式，而不在策略」。
 *
 * @param result  引擎解析结果
 * @param policy  解析质量策略
 * @param config  模块配置（用于判断对应能力是否被解析）
 */
export function runVerification(
  result: ParseResult,
  policy: ParsePolicy,
  config: ModuleConfig,
): VerificationReport {
  const checks: VerificationCheck[] = [];
  const counts = computeCounts(result);

  // --- 1. 前置：结构完整性 ---------------------------------------------- //
  checks.push(structuralCheck(result));

  // --- 2. 策略驱动的检查 ------------------------------------------------ //
  // 说明：若对应的能力开关被关闭，则该项本次并未解析，
  // 因此一律判为 skip（而不是 pass），避免「未解析 = 满足」的误判。

  // 加密：字段来自 ZIP 结构本身，不受能力开关影响。
  if (policy.allowEncrypted === undefined) {
    checks.push(
      skipped(VERIFICATION_CHECK_IDS.ENCRYPTION, 'Policy does not state a requirement for encryption.'),
    );
  } else if (result.encrypted === null) {
    checks.push(
      skipped(VERIFICATION_CHECK_IDS.ENCRYPTION, 'Encryption state could not be determined.'),
    );
  } else if (result.encrypted === false) {
    checks.push(check(VERIFICATION_CHECK_IDS.ENCRYPTION, 'pass', 'info', 'Artifact is not encrypted.'));
  } else if (policy.allowEncrypted) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.ENCRYPTION, 'pass', 'info', 'Artifact is encrypted, which the policy allows.'),
    );
  } else {
    checks.push(
      check(VERIFICATION_CHECK_IDS.ENCRYPTION, 'fail', 'error', 'Artifact is encrypted, which the policy forbids.'),
    );
  }

  // 标题
  if (!config.featureFlags.parseHeadings) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.HEADINGS, 'Headings were not parsed in this call (feature flag disabled).'));
  } else if (policy.requireHeadings === undefined) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.HEADINGS, 'Policy does not require headings.'));
  } else if (counts.headings > 0) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.HEADINGS, 'pass', 'info', 'Document contains headings.', {
        headings: counts.headings,
      }),
    );
  } else {
    checks.push(
      check(VERIFICATION_CHECK_IDS.HEADINGS, 'fail', 'error', 'Policy requires headings, but none were found.'),
    );
  }

  // 最少非空段落数
  if (!config.featureFlags.parseParagraphs) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.MIN_PARAGRAPHS, 'Paragraphs were not parsed in this call (feature flag disabled).'));
  } else if (policy.minParagraphs === undefined) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.MIN_PARAGRAPHS, 'Policy does not set a paragraph minimum.'));
  } else {
    const nonEmpty = result.blocks.reduce((total, block) => {
      if (block.kind !== 'paragraph') return total;
      return block.paragraph.text.trim() === '' ? total : total + 1;
    }, 0);
    const within = nonEmpty >= policy.minParagraphs;
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.MIN_PARAGRAPHS,
        within ? 'pass' : 'fail',
        within ? 'info' : 'error',
        `Document has ${nonEmpty} non-empty paragraph(s); policy requires at least ${policy.minParagraphs}.`,
        { nonEmpty, minimum: policy.minParagraphs },
      ),
    );
  }

  // 样式表存在性
  if (!config.featureFlags.parseStyles) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.STYLES, 'Styles were not parsed in this call (feature flag disabled).'));
  } else if (policy.requireStyles === undefined) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.STYLES, 'Policy does not require a style table.'));
  } else if (counts.styles > 0) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.STYLES, 'pass', 'info', 'Document declares a style table.', {
        styles: counts.styles,
      }),
    );
  } else {
    checks.push(
      check(VERIFICATION_CHECK_IDS.STYLES, 'fail', 'error', 'Policy requires a style table, but none was found.'),
    );
  }

  // 样式引用可解析
  if (!config.featureFlags.parseStyles) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.RESOLVED_STYLES, 'Styles were not parsed in this call (feature flag disabled).'));
  } else if (policy.requireResolvedStyles === undefined) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.RESOLVED_STYLES, 'Policy does not require resolved style references.'));
  } else if (hasIssue(result, 'STYLE_NOT_FOUND')) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.RESOLVED_STYLES, 'fail', 'error', 'Some paragraphs reference undefined styles.'),
    );
  } else {
    checks.push(
      check(VERIFICATION_CHECK_IDS.RESOLVED_STYLES, 'pass', 'info', 'All referenced styles are defined.'),
    );
  }

  // 无悬空关系
  if (!config.featureFlags.parseRelationships) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.DANGLING_RELATIONSHIPS, 'Relationships were not parsed in this call (feature flag disabled).'));
  } else if (policy.requireNoDanglingRelationships === undefined) {
    checks.push(skipped(VERIFICATION_CHECK_IDS.DANGLING_RELATIONSHIPS, 'Policy does not require intact relationship targets.'));
  } else if (hasIssue(result, 'DANGLING_RELATIONSHIP')) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.DANGLING_RELATIONSHIPS, 'fail', 'error', 'Some internal relationships target missing parts.'),
    );
  } else {
    checks.push(
      check(VERIFICATION_CHECK_IDS.DANGLING_RELATIONSHIPS, 'pass', 'info', 'All internal relationship targets exist.'),
    );
  }

  // --- 3. 收尾：序列化自检 ---------------------------------------------- //
  checks.push(serializableCheck(result, checks));

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
 * 用途：`inspect` 的快速失败策略使用它——warn 级失败不应阻止 inspect 返回结果，
 * 但 error 级失败（如结构不成立、策略明令禁止）应当阻止。
 */
export function hasBlockingFailure(report: VerificationReport): boolean {
  return report.checks.some((entry) => entry.status === 'fail' && entry.severity === 'error');
}

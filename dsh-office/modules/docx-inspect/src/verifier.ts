/**
 * 局部验证器 —— 在【本模块能力范围内】对产物与安全策略做检查。
 *
 * 职责边界（spec 第八节「Profile 集成要求」）：
 *   - 这里只做「结构 + 指标」层面的验证，属于 docx-inspect 的能力边界内；
 *   - 视觉/渲染类验证由 office-preview、共享验证套件由 office-test-kit 负责，
 *     那些不在本模块范围内；
 *   - 报告结构由本模块拥有，底层库无法改写。
 *
 * 三态语义（贯穿全部检查）：
 *   pass —— 策略声明了该要求，且产物满足；
 *   fail —— 策略声明了该要求，但产物不满足；
 *   skip —— 策略【未声明】该要求，因此无法判断。
 *
 * 之所以要区分 skip 与 pass：策略作者必须能表达「我确认允许」与
 * 「我根本没表态」这两种截然不同的意思，否则报告会产生虚假的安全感。
 */
import type { SafetyPolicy } from 'office-safety';

import type {
  JsonObject,
  ModuleConfig,
  VerificationCheck,
  VerificationReport,
  VerificationSummary,
} from './contract';
import {
  hasActiveX,
  hasAltChunks,
  hasDdeFields,
  hasEmbeddedObjects,
  hasMacros,
  type ProbeResult,
} from './domain/docx-inspect';

/** 检查项 id 常量表：稳定 id 是回归测试能做精确断言的前提。 */
export const VERIFICATION_CHECK_IDS = {
  STRUCTURE: 'structure.package',
  SERIALIZABLE: 'output.serializable',
  ENCRYPTION: 'policy.encryption',
  MACROS: 'policy.macros',
  EXTERNAL_LINKS: 'policy.externalLinks',
  EXTERNAL_TARGET_BUDGET: 'policy.externalTargetBudget',
  EMBEDDED_OBJECTS: 'policy.embeddedObjects',
  ALT_CHUNKS: 'policy.altChunks',
  ACTIVE_X: 'policy.activeX',
  DDE_FIELDS: 'policy.ddeFields',
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

/**
 * 策略布尔项的通用检查。
 *
 * 抽出这条 helper 是为了让「未声明即 skip」与「未探测即 skip」两种语义
 * 只写一次，避免各处实现出现分歧。
 *
 * @param probed 本次调用是否真的探测了该指标。若能力开关被关闭，则指标必然为空，
 *               此时【不能】判 pass（那会给出虚假的安全感），必须判 skip。
 */
function policyBooleanCheck(
  id: string,
  allowed: boolean | undefined,
  present: boolean,
  probed: boolean,
  presentMessage: string,
  absentMessage: string,
): VerificationCheck {
  if (!probed) {
    return check(
      id,
      'skip',
      'info',
      'Indicator was not probed in this call (feature flag disabled).',
    );
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

/**
 * 结构完整性检查。
 *
 * 这是唯一的「内部」检查（不受策略开关影响）：如果包结构本身不成立，
 * 那么任何策略结论都是不可靠的，必须无条件先确认。
 */
function structuralCheck(probe: ProbeResult): VerificationCheck {
  const problems: string[] = [];
  if (probe.container !== 'zip') {
    problems.push(`container is ${probe.container}`);
  }
  if (probe.documentKind !== 'wordprocessingml') {
    problems.push(`documentKind is ${probe.documentKind}`);
  }
  if (probe.partCount === 0) {
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
    { partCount: probe.partCount, xmlBackend: probe.xmlBackend },
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
function serializableCheck(probe: ProbeResult, report: Omit<VerificationReport, 'checks' | 'summary' | 'ok' | 'partial'>): VerificationCheck {
  try {
    // 用一个可被 JSON 序列化的代理对象来验证：若存在循环引用或
    // BigInt/函数等不可序列化值，这里会抛出。
    JSON.stringify({ probe, report });
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
 * @param probe   引擎探测结果
 * @param policy  安全策略（office-safety）
 * @param config  模块配置（用于判断是否启用指标探测）
 */
export function runVerification(
  probe: ProbeResult,
  policy: SafetyPolicy,
  config: ModuleConfig,
): VerificationReport {
  const checks: VerificationCheck[] = [];

  // --- 1. 前置：结构完整性 ---------------------------------------------- //
  checks.push(structuralCheck(probe));

  // --- 2. 策略驱动的指标检查 -------------------------------------------- //
  // 说明：若对应的能力开关被关闭，则该项指标本次并未被探测，
  // 因此一律判为 skip（而不是 pass），避免产生「未探测 = 安全」的误判。

  checks.push(
    policyBooleanCheck(
      VERIFICATION_CHECK_IDS.ENCRYPTION,
      policy.allowEncrypted,
      probe.encrypted === true,
      // 加密标记来自 ZIP 结构本身，不受能力开关影响，始终可用。
      true,
      'Artifact is encrypted.',
      'Artifact is not encrypted.',
    ),
  );

  // 宏：除 allowMacros 外，reviewOnMacros 表示「不禁止，但必须人工复核」，
  // 语义上等价于「禁止自动通过」→ 若携带宏则判 fail。
  const macrosProbed = config.featureFlags.detectMacros;
  if (!macrosProbed) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.MACROS,
        'skip',
        'info',
        'Indicator was not probed in this call (feature flag disabled).',
      ),
    );
  } else if (policy.allowMacros === undefined && policy.reviewOnMacros !== true) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.MACROS, 'skip', 'info', 'Policy does not state a requirement for this indicator.'),
    );
  } else if (!hasMacros(probe)) {
    checks.push(check(VERIFICATION_CHECK_IDS.MACROS, 'pass', 'info', 'Artifact carries no macros.'));
  } else if (policy.allowMacros === true && policy.reviewOnMacros !== true) {
    checks.push(
      check(VERIFICATION_CHECK_IDS.MACROS, 'pass', 'info', 'Artifact carries macros, which the policy allows.', {
        macroParts: probe.macroIndicators.map((indicator) => indicator.part).sort(),
      }),
    );
  } else {
    const reason = policy.reviewOnMacros === true ? 'requires manual review' : 'is not allowed';
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.MACROS,
        'fail',
        'error',
        `Artifact carries macros, which the policy ${reason}.`,
        {
          macroParts: probe.macroIndicators.map((indicator) => indicator.part).sort(),
          reviewOnMacros: policy.reviewOnMacros === true,
        },
      ),
    );
  }

  checks.push(
    policyBooleanCheck(
      VERIFICATION_CHECK_IDS.EXTERNAL_LINKS,
      policy.allowExternalLinks,
      probe.externalReferences.length > 0,
      config.featureFlags.detectExternalLinks,
      'Artifact declares external relationship targets.',
      'Artifact declares no external relationship targets.',
    ),
  );

  // 外部目标数量预算：仅当策略给了上限时才有意义。
  if (!config.featureFlags.detectExternalLinks) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.EXTERNAL_TARGET_BUDGET,
        'skip',
        'info',
        'Indicator was not probed in this call (feature flag disabled).',
      ),
    );
  } else if (policy.maxExternalTargets === undefined) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.EXTERNAL_TARGET_BUDGET,
        'skip',
        'info',
        'Policy does not set an external target budget.',
      ),
    );
  } else {
    const count = probe.externalReferences.length;
    const withinBudget = count <= policy.maxExternalTargets;
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.EXTERNAL_TARGET_BUDGET,
        withinBudget ? 'pass' : 'fail',
        withinBudget ? 'info' : 'error',
        `Artifact declares ${count} external target(s); policy budget is ${policy.maxExternalTargets}.`,
        { count, budget: policy.maxExternalTargets },
      ),
    );
  }

  checks.push(
    policyBooleanCheck(
      VERIFICATION_CHECK_IDS.EMBEDDED_OBJECTS,
      policy.allowEmbeddedObjects,
      hasEmbeddedObjects(probe),
      config.featureFlags.detectEmbeddedObjects,
      'Artifact embeds OLE/package objects.',
      'Artifact embeds no OLE/package objects.',
    ),
  );

  checks.push(
    policyBooleanCheck(
      VERIFICATION_CHECK_IDS.ALT_CHUNKS,
      policy.allowAltChunks,
      hasAltChunks(probe),
      config.featureFlags.detectAltChunks,
      'Artifact contains altChunk parts.',
      'Artifact contains no altChunk parts.',
    ),
  );

  // ActiveX 与 DDE 没有独立的策略开关：它们一律是「需要被注意」的指标，
  // 若存在则给出 warn 级 fail，让上层能按需升级为硬拒绝。
  if (!config.featureFlags.detectActiveX) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.ACTIVE_X,
        'skip',
        'info',
        'Indicator was not probed in this call (feature flag disabled).',
      ),
    );
  } else if (hasActiveX(probe)) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.ACTIVE_X,
        'fail',
        'warn',
        'Artifact contains ActiveX controls, which are never expected in a document pipeline.',
        { count: probe.embeddedObjects.filter((object) => object.kind === 'activeX').length },
      ),
    );
  } else {
    checks.push(check(VERIFICATION_CHECK_IDS.ACTIVE_X, 'pass', 'info', 'Artifact contains no ActiveX controls.'));
  }

  if (!config.featureFlags.detectDdeFields) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.DDE_FIELDS,
        'skip',
        'info',
        'Indicator was not probed in this call (feature flag disabled).',
      ),
    );
  } else if (hasDdeFields(probe)) {
    checks.push(
      check(
        VERIFICATION_CHECK_IDS.DDE_FIELDS,
        'fail',
        'warn',
        'Artifact contains DDE/DDEAUTO fields.',
        { count: probe.ddeFields.length },
      ),
    );
  } else {
    checks.push(check(VERIFICATION_CHECK_IDS.DDE_FIELDS, 'pass', 'info', 'Artifact contains no DDE fields.'));
  }

  // --- 3. 收尾：序列化自检 ---------------------------------------------- //
  const summary = summarize(checks);
  const pending = {
    policyId: policy.id,
    ok: summary.failed === 0,
    partial: summary.skipped > 0,
  };
  checks.push(serializableCheck(probe, pending));

  // 注意：serializable 检查本身也会改变统计，因此必须在加入它之后再汇总一次。
  const finalSummary = summarize(checks);
  return {
    ...pending,
    checks,
    summary: finalSummary,
    // 序列化检查若失败，整体结论必须随之更新。
    ok: finalSummary.failed === 0,
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
 * 用途：`inspect` 的快速失败策略使用它——warn 级失败（如 ActiveX 指标）
 * 不应阻止 inspect 返回结果，但 error 级失败（如结构不成立、策略明令禁止）应当阻止。
 */
export function hasBlockingFailure(report: VerificationReport): boolean {
  return report.checks.some((entry) => entry.status === 'fail' && entry.severity === 'error');
}

/**
 * docx-parse —— Profile 注册入口。
 *
 * 本文件是模块的「composition root」：把引擎、映射器、验证器装配成三个公开 handler，
 * 并对外交出与 `module.json` 一一对应的模块清单。所有编排都止步于此——
 * 模块不感知 Profile 如何调度它，也不感知谁在调用它。
 *
 * 三层职责在这里被明确切分：
 *   - 引擎（adapter.ts）   ：拿字节，产出中立的 `ParseResult`；
 *   - 映射（mapper.ts）    ：裁决 + 派生 + 归一化，产出公开 IR / FormatProfile；
 *   - 验证（verifier.ts）  ：对 `ParseResult` 做局部质量检查。
 * index.ts 只做「顺序编排 + 错误归一化 + 信封组装」，不含任何解析逻辑。
 */
import type { ArtifactRef } from 'office-core';

import {
  DOCX_PARSE_CAPABILITIES,
  DOCX_PARSE_MODULE_ID,
  type DocxExecuteFn,
  type DocxModuleInput,
  type DocxModuleOutput,
  type DocxParseFn,
  type DocxParseHandlers,
  type DocxParseModule,
  type DocxParseModuleDefinition,
  type ExecuteOutput,
  type InspectOutput,
  type ModuleConfig,
  type ModuleContext,
  type ParsePolicy,
  type ProfileRegistry,
  type TelemetrySummary,
  type VerifyOutput,
  type DocxVerifyFn,
} from './contract';
import { configSchema, withCallOverrides, type ConfigOverrides } from './config';
import { computeCounts, type ParseResult } from './domain/docx-parse';
import {
  PythonDocxEngine,
  createDeepEngine,
  resolveArtifactPath,
  type DocxEngine,
} from './engine/adapter';
import { DocxParseError, toDocxParseError } from './errors';
import {
  assertParseUsable,
  refineArtifactRef,
  toDocxParseIR,
  toFormatProfile,
  toWarnings,
} from './mapper';
import { hasBlockingFailure, runVerification } from './verifier';
import { createTelemetryLogger, noopTelemetry } from './telemetry';

/** 模块版本；与 `module.json` 保持同步。 */
export const DOCX_PARSE_VERSION = '0.2.0';

/**
 * 模块清单。
 *
 * 与 `module.json` 是「同一份事实的两种呈现」：JSON 供 Profile registry 静态读取，
 * 这里的类型化常量供代码内消费。两者必须一致，因此回归测试会逐字段比对。
 */
export const DOCX_PARSE_DEFINITION: DocxParseModuleDefinition = {
  id: DOCX_PARSE_MODULE_ID,
  version: DOCX_PARSE_VERSION,
  profileGroup: 'DOCX',
  summary:
    'Structural parsing of DOCX: paragraphs, headings, tables, styles, relationships, comments, footnotes mapped to a normalized IR.',
  capabilities: DOCX_PARSE_CAPABILITIES,
  dependencies: [
    { name: 'office-core', kind: 'module' },
    { name: 'office-safety', kind: 'module' },
    { name: 'office-files', kind: 'module', optional: true },
    { name: 'office-test-kit', kind: 'module', optional: true },
    { name: 'python', kind: 'runtime' },
  ],
  configSchema: configSchema(),
};

/** 创建模块时的可选注入点（用于单测替换引擎或遥测）。 */
export interface CreateDocxParseModuleOptions {
  /** 覆盖默认的轻量引擎实现；省略时使用内置 Python 桥接。`inspect` 恒使用它。 */
  engine?: DocxEngine;
  /**
   * 覆盖 `execute` / `verify` 所用的深度引擎。
   *
   * 不传时按 `config.engine.deep` 自动构造；`config.engine.deep` 也未配置时，
   * 这两条路径回退到轻量引擎，行为与未引入深度引擎时完全一致。
   */
  deepEngine?: DocxEngine;
  /** 覆盖遥测接收端。 */
  telemetry?: ModuleContext['telemetry'];
  /** 覆盖日志实现。 */
  logger?: ModuleContext['logger'];
}

/** 一次解析流水线的内部产出。 */
interface ParseOutcome {
  result: ParseResult;
  /** 应用了单次调用覆盖之后、本次请求实际生效的配置。 */
  config: ModuleConfig;
  engineMs: number;
  startedAt: number;
  /** 本次实际执行的引擎名，用于遥测与诊断（区分轻量/深度路径）。 */
  engineName: string;
}

/**
 * 校验输入信封。
 *
 * spec 要求「输入必须可序列化」，且 `artifactRef` 至少要有可解析的 uri。
 * 这里只做形状校验，uri 的 scheme 合法性交给 `resolveArtifactPath`。
 *
 * @throws DocxParseError(INVALID_INPUT)
 */
function validateInput(input: DocxModuleInput): void {
  if (typeof input.requestId !== 'string' || input.requestId.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'requestId must be a non-empty string');
  }
  const ref: ArtifactRef | undefined = input.artifactRef;
  if (ref === null || typeof ref !== 'object' || typeof ref.uri !== 'string' || ref.uri.trim() === '') {
    throw new DocxParseError('INVALID_INPUT', 'artifactRef.uri must be a non-empty string');
  }
  // 输入必须能跨进程传输：不可序列化即拒绝，避免问题推迟到引擎边界才暴露。
  try {
    JSON.stringify(input);
  } catch (error) {
    throw new DocxParseError('INVALID_INPUT', 'Input is not JSON-serializable', {
      cause: error,
    });
  }
}

/** 把单次调用的选项翻译成配置覆盖项。 */
function overridesFromOptions(input: DocxModuleInput): ConfigOverrides {
  const options = input.options;
  if (options === undefined) return {};
  const overrides: ConfigOverrides = {};
  if (options.limits !== undefined) overrides.limits = options.limits;
  if (options.featureFlags !== undefined) overrides.featureFlags = options.featureFlags;
  if (options.timeoutMs !== undefined) overrides.timeoutMs = options.timeoutMs;
  return overrides;
}

/** 由解析结果计算遥测摘要（不含任何文档内容）。 */
function buildTelemetrySummary(
  result: ParseResult,
  engineName: string,
  engineMs: number,
  startedAt: number,
): TelemetrySummary {
  const counts = computeCounts(result);
  return {
    engine: engineName,
    parseVersion: result.parseVersion,
    xmlBackend: result.xmlBackend,
    engineMs,
    totalMs: Date.now() - startedAt,
    blockCount: counts.blocks,
    tableCount: counts.tables,
    styleCount: counts.styles,
  };
}

/** 组装输出信封的公共部分。 */
function envelope<TResult>(
  input: DocxModuleInput,
  result: TResult,
  warnings: ReturnType<typeof toWarnings>,
  telemetry: TelemetrySummary,
  verification?: ReturnType<typeof runVerification>,
): DocxModuleOutput<TResult> {
  const output: DocxModuleOutput<TResult> = {
    moduleId: DOCX_PARSE_MODULE_ID,
    requestId: input.requestId,
    operation: input.operation,
    result,
    // 解析不产出新文件：这里恒为空数组，保持信封形状稳定。
    artifacts: [],
    warnings,
    telemetry,
  };
  if (verification !== undefined) {
    output.verification = verification;
  }
  return output;
}

/**
 * 创建 docx-parse 模块实例。
 *
 * @param context Profile 注入的运行时上下文（配置 + 可选遥测/日志）
 * @param options 测试用注入点
 */
export function createDocxParseModule(
  context: ModuleContext,
  options: CreateDocxParseModuleOptions = {},
): DocxParseModule {
  const telemetry = options.telemetry ?? context.telemetry ?? noopTelemetry;
  const logger = options.logger ?? context.logger ?? createTelemetryLogger(telemetry);
  // 轻量引擎：inspect 无条件使用它，保证「先看一眼」始终是秒级操作。
  const engine = options.engine ?? new PythonDocxEngine(context.config.engine, telemetry);
  // 深度引擎：只在配置声明时存在。为 null 时 execute/verify 回退到轻量引擎，
  // 因此「启用 Docling」是一个纯增量开关，不会改变默认行为。
  const deepEngine =
    options.deepEngine ?? createDeepEngine(context.config.engine, telemetry);
  const deepTimeoutMs = context.config.engine.deep?.timeoutMs;

  /**
   * 公共流水线：校验 → 解析 → 返回中立结果。
   *
   * 三个 handler 共享它，从而保证「解析与裁决」在所有接口上完全一致；
   * 差异只体现在「用哪个引擎」与「裁决之后怎么组装输出」。
   *
   * @param engineForCall 本次调用使用的引擎；深度引擎只可能由 execute/verify 传入。
   */
  async function runPipeline(
    input: DocxModuleInput,
    engineForCall: DocxEngine,
  ): Promise<ParseOutcome> {
    validateInput(input);
    const effective = withCallOverrides(context.config, overridesFromOptions(input));
    const artifactPath = resolveArtifactPath(input.artifactRef.uri);

    // 深度引擎需要更宽的超时预算（模型冷启动）；未配置时沿用调用级超时。
    const usingDeepEngine = engineForCall !== engine;
    const timeoutMs =
      usingDeepEngine && deepTimeoutMs !== undefined ? deepTimeoutMs : effective.timeoutMs;

    const startedAt = Date.now();
    const result = await engineForCall.parse({
      artifactPath,
      limits: effective.limits,
      featureFlags: effective.featureFlags,
      timeoutMs,
    });

    return {
      result,
      config: effective,
      engineMs: Date.now() - startedAt,
      startedAt,
      engineName: engineForCall.name,
    };
  }

  /** inspect：只回答「这是什么产物」，并按需附带策略检查。恒走轻量引擎。 */
  const inspect: DocxParseFn = async (input) => {
    try {
      const { result, config, engineMs, startedAt, engineName } = await runPipeline(
        input,
        engine,
      );
      // 裁决先于映射：失败绝不产出「半成品」的输出。
      assertParseUsable(result, config, input.options);
      const warnings = toWarnings(result, input.options, config);
      const profile = toFormatProfile(result, input.options);

      let verification;
      const shouldVerify = input.policy !== undefined || input.options?.verifyAfterInspect === true;
      if (shouldVerify) {
        const policy: ParsePolicy = input.policy ?? { id: `${DOCX_PARSE_MODULE_ID}.inspect` };
        verification = runVerification(result, policy, config);
        // inspect 的失败语义之一就是「安全策略拒绝」：加密策略是唯一会在此硬拒的项。
        if (policy.allowEncrypted === false && result.encrypted === true) {
          throw new DocxParseError(
            'SAFETY_POLICY_DENIED',
            'Artifact is encrypted, which the supplied policy forbids',
            { details: { policyId: policy.id, encrypted: true } },
          );
        }
        // 结构级硬失败同样视为策略拒绝（正常情况下 assertParseUsable 已拦截）。
        if (hasBlockingFailure(verification)) {
          throw new DocxParseError(
            'SAFETY_POLICY_DENIED',
            'Artifact does not satisfy the supplied policy',
            {
              details: {
                policyId: policy.id,
                failed: verification.checks
                  .filter((entry) => entry.status === 'fail')
                  .map((entry) => entry.id),
              },
            },
          );
        }
      }

      logger.info('docx-parse.inspect.completed', {
        format: profile.format,
        blocks: computeCounts(result).blocks,
      });

      const output: InspectOutput = envelope(
        input,
        profile,
        warnings,
        buildTelemetrySummary(result, engineName, engineMs, startedAt),
        verification,
      );
      return output;
    } catch (error) {
      throw toDocxParseError(error, 'PARSE_FAILED');
    }
  };

  /** execute：产出归一化 IR 与补全后的 artifact 引用。配置了深度引擎时改走它。 */
  const execute: DocxExecuteFn = async (input) => {
    try {
      const { result, config, engineMs, startedAt, engineName } = await runPipeline(
        input,
        deepEngine ?? engine,
      );
      assertParseUsable(result, config, input.options);
      const warnings = toWarnings(result, input.options, config);
      const ir = toDocxParseIR(result, input.options);
      const artifact = refineArtifactRef(input.artifactRef, result);

      logger.info('docx-parse.execute.completed', {
        blocks: ir.counts.blocks,
        tables: ir.counts.tables,
        styles: ir.counts.styles,
      });

      const output: ExecuteOutput = envelope(
        input,
        { ir, artifact },
        warnings,
        buildTelemetrySummary(result, engineName, engineMs, startedAt),
      );
      return output;
    } catch (error) {
      throw toDocxParseError(error, 'PARSE_FAILED');
    }
  };

  /** verify：产出局部验证报告。配置了深度引擎时改走它。 */
  const verify: DocxVerifyFn = async (input) => {
    try {
      const policy = input.policy;
      if (policy === null || typeof policy !== 'object' || typeof policy.id !== 'string' || policy.id.trim() === '') {
        throw new DocxParseError('INVALID_INPUT', 'A policy with a non-empty id is required for verify');
      }

      const { result, config, engineMs, startedAt, engineName } = await runPipeline(
        input,
        deepEngine ?? engine,
      );
      assertParseUsable(result, config, input.options);
      const warnings = toWarnings(result, input.options, config);

      // 检查项的 fail 通过 `report.ok = false` 表达，而不是抛错——
      // 「文档不满足策略」是一个有效结论，不是调用失败。
      // VERIFICATION_FAILED 只保留给「报告本身无法产出」的情形。
      let report;
      try {
        report = runVerification(result, policy, config);
      } catch (error) {
        throw new DocxParseError('VERIFICATION_FAILED', 'Could not produce a verification report', {
          cause: error,
        });
      }

      logger.info('docx-parse.verify.completed', {
        ok: report.ok,
        failed: report.summary.failed,
        skipped: report.summary.skipped,
      });

      const output: VerifyOutput = envelope(
        input,
        report,
        warnings,
        buildTelemetrySummary(result, engineName, engineMs, startedAt),
      );
      return output;
    } catch (error) {
      throw toDocxParseError(error, 'VERIFICATION_FAILED');
    }
  };

  const handlers: DocxParseHandlers = { inspect, execute, verify };

  return {
    definition: DOCX_PARSE_DEFINITION,
    handlers,
    async dispose(): Promise<void> {
      // 两个引擎各持有自己的资源；深度引擎可能为 null（未配置）。
      await Promise.all([engine.dispose(), deepEngine?.dispose()]);
    },
  };
}

/**
 * 把模块注册进 Profile。
 *
 * registry 端口刻意做得极窄：本模块只负责「把自己交出去」，
 * 至于生命周期、依赖注入顺序、配置覆盖，全部由 Profile 决定。
 */
export async function register(
  registry: ProfileRegistry,
  context: ModuleContext,
  options: CreateDocxParseModuleOptions = {},
): Promise<DocxParseModule> {
  const module = createDocxParseModule(context, options);
  await registry.registerModule(module);
  return module;
}

/* -------------------------------------------------------------------------- */
/* 对外导出                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * 公开契约到此为止。
 *
 * `package.json` 的 `exports` 只映射了 `"."` → 本文件，所以「本文件导出了什么」
 * 就等于「调用方能依赖什么」。据此只导出四类东西：
 *
 *   1. 身份 —— 模块清单、能力表、错误码/告警码表（与 module.json 一一对应）；
 *   2. 装配 —— createDocxParseModule / register 及其选项类型；
 *   3. 配置 —— resolveConfig / configSchema 与 ConfigOverrides；
 *   4. 扩展 —— DocxEngine 端口、错误类型，以及全部输入输出类型。
 *
 * 映射、领域派生、验证、遥测与具体引擎实现都是内部接缝，见 `./internal.ts`。
 * 它们不构成契约，可以自由重构而不属于破坏性变更。
 */

export { configSchema, resolveConfig } from './config';
export type { ConfigOverrides } from './config';

// 引擎端口是公开的扩展点：调用方可以注入自己的实现（见 CreateDocxParseModuleOptions）。
// 具体实现（PythonDocxEngine / DoclingEngine 与脚本路径解析）属于内部，见 ./internal.ts。
export type { DocxEngine, ParseRequest } from './engine/adapter';

export { DocxParseError, isDocxParseError, toDocxParseError } from './errors';

export {
  DOCX_PARSE_CAPABILITIES,
  DOCX_PARSE_ERROR_CODES,
  DOCX_PARSE_MODULE_ID,
  DOCX_PARSE_WARNING_CODES,
} from './contract';

export type {
  DeepEngineConfig,
  DocxExecuteFn,
  DocxModuleInput,
  DocxModuleOutput,
  DocxOperation,
  DocxParseCapability,
  DocxParseErrorCode,
  DocxParseFn,
  DocxParseHandlers,
  DocxParseModule,
  DocxParseModuleDefinition,
  DocxParseOptions,
  DocxParseWarningCode,
  DocxVerifyFn,
  EngineConfig,
  ExecuteInput,
  ExecuteOutput,
  FeatureFlags,
  InspectInput,
  InspectOutput,
  JsonObject,
  JsonValue,
  LimitConfig,
  Logger,
  ModuleConfig,
  ModuleContext,
  ModuleDependencyRef,
  ParsePolicy,
  ProfileRegistry,
  TelemetryEvent,
  TelemetryLevel,
  TelemetrySink,
  TelemetrySummary,
  VerificationCheck,
  VerificationReport,
  VerificationSummary,
  VerifyInput,
  VerifyOutput,
} from './contract';

export type {
  Alignment,
  Block,
  CommentRecord,
  ContainerKind,
  DocumentCounts,
  DocumentKind,
  DocumentMetadata,
  DocumentVariant,
  DocxParseIR,
  FootnoteRecord,
  OutlineEntry,
  ParagraphRecord,
  ParseIssue,
  ParseResult,
  RelationshipRecord,
  RunRecord,
  StyleRecord,
  StyleType,
  TableCell,
  TableRecord,
} from './domain/docx-parse';

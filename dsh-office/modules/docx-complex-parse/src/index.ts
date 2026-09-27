/**
 * docx-complex-parse 的 Profile 注册入口（装配层）。
 *
 * 本文件只做三件事：
 *   1. 把配置、引擎、mapper、verifier 拼装成一个可注册的模块；
 *   2. 实现 `inspect` / `execute` / `verify` 三个 handler；
 *   3. 把内部异常统一收敛为本模块的错误分类。
 *
 * 边界纪律：
 *   - 对外只暴露 `contract.ts` 定义的接口，不泄漏引擎与 mapper 的内部形态；
 *   - 不实现任何上层编排（依赖解析、加载顺序、生命周期由 Profile 决定）；
 *   - 可以在 Profile 全量启动【之前】单独加载并测试。
 */
import type { ArtifactRef } from 'office-core';
import type {
  DocxComplexExecuteFn,
  DocxComplexInspectFn,
  DocxComplexParseHandlers,
  DocxComplexParseModule,
  DocxComplexParseModuleDefinition,
  DocxComplexParseOptions,
  DocxComplexVerifyFn,
  ExecuteOutput,
  InspectOutput,
  ModuleConfig,
  ModuleContext,
  ModuleDependencyRef,
  TelemetrySink,
  TelemetrySummary,
  VerificationReport,
  VerifyOutput,
} from './contract';
import { DOCX_COMPLEX_PARSE_CAPABILITIES, DOCX_COMPLEX_PARSE_MODULE_ID } from './contract';
import type { ConfigOverrides } from './config';
import { configSchema, resolveConfig, withCallOverrides } from './config';
import { DocxComplexParseError, toDocxComplexParseError } from './errors';
import {
  assertDeclaredTypeMatches,
  assertParseUsable,
  collectLowConfidence,
  countTableCells,
  extensionOf,
  refineArtifactRef,
  toComplexContent,
  toComplexIR,
  toFormatProfile,
  toWarnings,
  type BudgetOverrun,
  type MapperWarning,
} from './mapper';
import { createTelemetryLogger } from './telemetry';
import type { DocxEngine } from './engine/adapter';
import { createEngine } from './engine/adapter';
import type { RawParseResult } from './domain/docx-complex-parse';
import { parseRawResult } from './domain/docx-complex-parse';
import type { Warning } from 'office-core';
import { hasBlockingFailure, verify as runVerify } from './verifier';

/** 模块版本。必须与 `module.json` 中的 `version` 保持一致。 */
export const DOCX_COMPLEX_PARSE_VERSION = '0.1.0';

/** 依赖声明：module 指 Profile 内模块，runtime 指外部运行时。 */
const DEPENDENCIES: readonly ModuleDependencyRef[] = [
  { name: 'office-core', kind: 'module' },
  { name: 'office-safety', kind: 'module' },
  { name: 'office-files', kind: 'module', optional: true },
  { name: 'office-test-kit', kind: 'module', optional: true },
  // rdocx 是本模块的必需依赖：没有它就只剩结构线，来源坐标无从产出。
  { name: 'python', kind: 'runtime' },
];

/** 模块清单（与 `module.json` 一一对应）。 */
export const DOCX_COMPLEX_PARSE_DEFINITION: DocxComplexParseModuleDefinition = {
  id: DOCX_COMPLEX_PARSE_MODULE_ID,
  version: DOCX_COMPLEX_PARSE_VERSION,
  profileGroup: 'DOCX',
  summary:
    'Complex DOCX parsing: logical table grids, page relationships, floating objects, ' +
    'source coordinates, and per-finding confidence.',
  capabilities: DOCX_COMPLEX_PARSE_CAPABILITIES,
  dependencies: DEPENDENCIES,
  configSchema: configSchema(),
};

/** 创建模块时的可选注入点。 */
export interface CreateModuleOptions {
  /** 配置覆盖；未提供的字段使用 `config.ts` 的默认值。 */
  config?: ConfigOverrides;
  /** 注入自定义引擎（测试替身或未来替换的实现）。 */
  engine?: DocxEngine;
  /** 遥测接收端。 */
  telemetry?: TelemetrySink;
}

/**
 * 可注册的模块实例。
 *
 * 在契约的 `{ definition, handlers }` 之上多一个 `dispose`：引擎可能持有进程、
 * 连接或缓存，Profile 关闭时需要有机会释放。作为扩展属性提供，因此按契约类型
 * 消费的调用方不受影响。
 */
export interface DocxComplexParseModuleInstance extends DocxComplexParseModule {
  /** 释放引擎资源。可重复调用。 */
  dispose(): Promise<void>;
}

/** 一次调用的共同产物：原始观察 + 生效配置，供三个 handler 复用。 */
interface PreparedCall {
  active: ModuleConfig;
  raw: RawParseResult;
  engineMs: number;
  /** 被突破的预算；限额关闭时非空，调用方应把它们降级成告警。 */
  overruns: BudgetOverrun[];
}

/**
 * 构建一个可注册的模块实例。
 *
 * 【纯函数式工厂】：不做 I/O、不读全局配置、不注册副作用，
 * 因此测试里可以创建多个互不干扰的实例。
 */
export function createDocxComplexParseModule(
  options: CreateModuleOptions = {},
): DocxComplexParseModuleInstance {
  const config: ModuleConfig = resolveConfig(options.config);
  const context: ModuleContext = { config };
  if (options.telemetry !== undefined) {
    context.telemetry = options.telemetry;
  }
  context.logger = createTelemetryLogger(options.telemetry);

  // 引擎可替换：默认是 Python + rdocx 桥接，测试或未来实现可注入其他后端。
  const engine: DocxEngine = options.engine ?? createEngine(config.engine, options.telemetry);

  /** 按单次调用的 options 计算实际生效的配置。 */
  const effectiveConfig = (opts: DocxComplexParseOptions | undefined): ModuleConfig => {
    const overrides: ConfigOverrides = {};
    if (opts?.limits !== undefined) overrides.limits = opts.limits;
    if (opts?.featureFlags !== undefined) overrides.featureFlags = opts.featureFlags;
    if (opts?.timeoutMs !== undefined) overrides.timeoutMs = opts.timeoutMs;
    if (opts?.confidenceFloor !== undefined) overrides.confidenceFloor = opts.confidenceFloor;
    // 无覆盖时直接复用已冻结的配置，避免无谓重建。
    if (Object.keys(overrides).length === 0) return config;
    return withCallOverrides(config, overrides);
  };

  /**
   * 跑引擎、做类型核对与可用性判定。
   *
   * 三个 handler 共用这一段，保证「同一份 artifact 在 inspect / execute / verify
   * 下的可用性结论完全一致」——否则调用方会看到互相矛盾的结果。
   */
  const prepare = async (
    input: { artifactRef: ArtifactRef; options?: DocxComplexParseOptions | undefined },
  ): Promise<PreparedCall> => {
    const active = effectiveConfig(input.options);
    const startedAt = Date.now();

    let raw: RawParseResult;
    try {
      raw = parseRawResult(await engine.parse({
        // 传原始 uri：路径解析（含 unsupported scheme 的拒绝）是 adapter 的职责，
        // 在这里再做一次会让「换引擎」时行为不一致。
        artifactPath: input.artifactRef.uri,
        limits: active.limits,
        featureFlags: active.featureFlags,
        timeoutMs: active.timeoutMs,
      }));
    } catch (error) {
      // 引擎层抛出的已是模块错误（adapter 负责翻译），这里只兜底非预期异常。
      throw toDocxComplexParseError(error, 'ENGINE_FAILED', {
        message: 'Parse engine failed',
        path: input.artifactRef.uri,
      });
    }

    // 声明 vs 内容的核对先于可用性判定：伪装成 .docx 的 .docm 应当被明确拒绝，
    // 而不是被当成一份「格式不认识」的文档。
    assertDeclaredTypeMatches(raw, {
      extension: input.options?.declaredExtension,
      mimeType: input.options?.declaredMimeType,
    });
    const overruns = assertParseUsable(raw, active);

    return { active, raw, engineMs: Date.now() - startedAt, overruns };
  };

  /**
   * 预算超支在限额关闭时降级为告警。
   *
   * 之所以不在这里抛错：`assertParseUsable` 已经按 `enforceLimits` 决定过抛不抛，
   * 走到这一步说明配置选择「只记录、不拦截」。静默放过是最糟的选项——
   * 调用方必须知道这份结果超出了自己声明的预算。
   */
  const overrunWarnings = (overruns: readonly BudgetOverrun[]): MapperWarning[] =>
    overruns.map((overrun) => ({
      code: 'LIMIT_APPLIED',
      message: `${overrun.limit} was exceeded (${overrun.value} > ${overrun.max}) but enforcement is disabled.`,
    }));

  /** 汇总遥测摘要（不含文档内容与绝对路径）。 */
  const buildTelemetry = (
    raw: RawParseResult,
    active: ModuleConfig,
    counts: { blocks: number; tables: number; cells: number; floats: number },
    startedAt: number,
    engineMs: number,
  ): TelemetrySummary => ({
    parseVersion: raw.parseVersion,
    engineDriver: active.engine.driver,
    engineMs,
    totalMs: Date.now() - startedAt,
    partCount: raw.parts.length,
    relationshipCount: raw.relationships.length,
    blockCount: counts.blocks,
    tableCount: counts.tables,
    floatingObjectCount: counts.floats,
    pageCount: raw.layout.pages.length,
    pageGeometryAvailable: raw.layout.available,
  });

  // ---------------------------------------------------------------------- //
  // inspect：回答「这是什么产物」                                             //
  // ---------------------------------------------------------------------- //
  const inspect: DocxComplexInspectFn = async (input) => {
    const startedAt = Date.now();
    try {
      const { active, raw, engineMs, overruns } = await prepare(input);
      const profile = toFormatProfile(raw, { extension: extensionOf(input.artifactRef) });

      const warnings: Warning[] = toWarnings([...raw.warnings, ...overrunWarnings(overruns)]);
      // inspect 不产出内容，因此没有网格告警可报；但「没有内容模型」这件事
      // 必须说清楚，否则调用方会以为 inspect 已经覆盖了结构。
      if (raw.indicators.macros.length > 0) {
        warnings.push({
          code: 'UNSUPPORTED_CONTENT',
          message:
            `The package carries ${raw.indicators.macros.length} macro indicator(s); ` +
            'this module never executes them but also does not analyse their content.',
          severity: 'warn',
        });
      }

      let verification: VerificationReport | undefined;
      if (input.policy !== undefined) {
        // 调用方给了策略：先跑一次最小执行（要判断宏/预算就得先看到内容），
        // 硬失败时明确拒绝，而不是交给调用方自己去比对 profile。
        const { content } = toComplexContent(raw, active.featureFlags);
        verification = runVerify({
          ir: toComplexIR(raw, content, profile),
          policy: input.policy,
          confidenceFloor: active.confidenceFloor,
        });
        if (hasBlockingFailure(verification)) {
          throw new DocxComplexParseError(
            'SAFETY_POLICY_DENIED',
            'Artifact is rejected by the supplied safety policy.',
            {
              details: {
                policyId: input.policy.id,
                failedChecks: verification.checks
                  .filter((entry) => entry.status === 'fail')
                  .map((entry) => entry.id),
              },
            },
          );
        }
      } else if (input.options?.verifyAfterInspect === true) {
        // 仅做「结构自检」：用一个不声明任何要求的空策略，因此结果只会是
        // pass 或 skip，不会因为策略而拒绝文档。
        const { content } = toComplexContent(raw, active.featureFlags);
        verification = runVerify({
          ir: toComplexIR(raw, content, profile),
          policy: { id: 'inspect.self-check' },
          confidenceFloor: active.confidenceFloor,
        });
      }

      const output: InspectOutput = {
        moduleId: DOCX_COMPLEX_PARSE_MODULE_ID,
        requestId: input.requestId,
        operation: 'inspect',
        result: profile,
        // 解析不产生新文件，因此 artifacts 恒为空数组（保持形状一致）。
        artifacts: [],
        warnings,
        telemetry: buildTelemetry(
          raw,
          active,
          { blocks: raw.blocks.length, tables: raw.blocks.filter((block) => block.kind === 'table').length,
            cells: 0, floats: raw.floats?.length ?? 0 },
          startedAt,
          engineMs,
        ),
      };
      if (verification !== undefined) output.verification = verification;
      return output;
    } catch (error) {
      context.logger?.error('docx-complex-parse.inspect.failed', { requestId: input.requestId });
      throw toDocxComplexParseError(error, 'PARSE_FAILED', { path: input.artifactRef.uri });
    }
  };

  // ---------------------------------------------------------------------- //
  // execute：产出复杂结构 IR                                                  //
  // ---------------------------------------------------------------------- //
  const execute: DocxComplexExecuteFn = async (input) => {
    const startedAt = Date.now();
    try {
      const { active, raw, engineMs, overruns } = await prepare(input);
      const profile = toFormatProfile(raw, { extension: extensionOf(input.artifactRef) });

      const mapped = toComplexContent(raw, active.featureFlags);
      // 三类告警合流：引擎/mapper 的观察 → 预算超支 → 低置信度结论。
      const warnings = toWarnings([
        ...mapped.warnings,
        ...overrunWarnings(overruns),
        ...collectLowConfidence(mapped.content, active.confidenceFloor),
      ]);

      const tables = mapped.content.blocks.filter((block) => block.kind === 'table').length;

      const output: ExecuteOutput = {
        moduleId: DOCX_COMPLEX_PARSE_MODULE_ID,
        requestId: input.requestId,
        operation: 'execute',
        result: {
          ir: toComplexIR(raw, mapped.content, profile),
          artifact: refineArtifactRef(input.artifactRef, raw),
        },
        artifacts: [],
        warnings,
        telemetry: buildTelemetry(
          raw,
          active,
          {
            blocks: mapped.content.blocks.length,
            tables,
            cells: countTableCells(mapped.content),
            floats: mapped.content.floats.length,
          },
          startedAt,
          engineMs,
        ),
      };
      return output;
    } catch (error) {
      context.logger?.error('docx-complex-parse.execute.failed', { requestId: input.requestId });
      throw toDocxComplexParseError(error, 'PARSE_FAILED', { path: input.artifactRef.uri });
    }
  };

  // ---------------------------------------------------------------------- //
  // verify：按策略评估并产出局部验证报告                                      //
  // ---------------------------------------------------------------------- //
  const verify: DocxComplexVerifyFn = async (input) => {
    const startedAt = Date.now();
    try {
      const { active, raw, engineMs, overruns } = await prepare(input);
      const profile = toFormatProfile(raw, { extension: extensionOf(input.artifactRef) });

      const mapped = toComplexContent(raw, active.featureFlags);
      const report = runVerify({
        ir: toComplexIR(raw, mapped.content, profile),
        policy: input.policy,
        confidenceFloor: active.confidenceFloor,
      });

      const warnings = toWarnings([
        ...mapped.warnings,
        ...overrunWarnings(overruns),
        ...collectLowConfidence(mapped.content, active.confidenceFloor),
      ]);
      // 报告结论与告警互为补充：把「验证未通过」也提升为一条告警，
      // 让只消费 warnings 的上层同样能感知到。
      if (!report.ok) {
        warnings.push({
          code: 'VERIFICATION_FAILED',
          message: 'Artifact did not satisfy every declared policy requirement.',
          severity: 'error',
          details: { policyId: report.policyId ?? null, failed: report.summary.failed },
        });
      }

      const output: VerifyOutput = {
        moduleId: DOCX_COMPLEX_PARSE_MODULE_ID,
        requestId: input.requestId,
        operation: 'verify',
        result: report,
        artifacts: [],
        warnings,
        verification: report,
        telemetry: buildTelemetry(
          raw,
          active,
          {
            blocks: mapped.content.blocks.length,
            tables: mapped.content.blocks.filter((block) => block.kind === 'table').length,
            cells: countTableCells(mapped.content),
            floats: mapped.content.floats.length,
          },
          startedAt,
          engineMs,
        ),
      };
      return output;
    } catch (error) {
      context.logger?.error('docx-complex-parse.verify.failed', { requestId: input.requestId });
      throw toDocxComplexParseError(error, 'VERIFICATION_FAILED', { path: input.artifactRef.uri });
    }
  };

  const handlers: DocxComplexParseHandlers = { inspect, execute, verify };

  return {
    definition: DOCX_COMPLEX_PARSE_DEFINITION,
    handlers,
    dispose: async () => {
      await engine.dispose();
    },
  };
}

/**
 * 便捷函数：构建模块并注册到 Profile registry。
 *
 * 只调用 registry 的 `registerModule`，不做任何编排——加载顺序、依赖解析、
 * 生命周期均由 Profile 决定。
 */
export async function register(
  registry: ProfileRegistryLike,
  options: CreateModuleOptions = {},
): Promise<DocxComplexParseModuleInstance> {
  const module = createDocxComplexParseModule(options);
  await registry.registerModule(module);
  return module;
}

/**
 * Profile registry 的最小端口。
 *
 * 只声明本模块真正用到的一个方法，从而在缺少真实 Profile 包时也能独立编译与测试。
 */
export interface ProfileRegistryLike {
  registerModule(module: DocxComplexParseModule): void | Promise<void>;
}

/* -------------------------------------------------------------------------- */
/* 对外导出                                                                     */
/* -------------------------------------------------------------------------- */

export * from './contract';
export type { ConfigOverrides } from './config';
export {
  configSchema,
  DEFAULT_CONFIDENCE_FLOOR,
  DEFAULT_ENGINE,
  DEFAULT_FEATURE_FLAGS,
  DEFAULT_LIMITS,
  DEFAULT_TIMEOUT_MS,
  resolveConfig,
  toConfigOverrides,
  withCallOverrides,
} from './config';
export { DocxComplexParseError, isDocxComplexParseError, toDocxComplexParseError } from './errors';
export type { TelemetryEvent, TelemetryLevel } from './contract';
export { createTelemetryLogger, InMemoryTelemetry, noopTelemetry, withSpan } from './telemetry';
export type { DocxEngine, ParseRequest } from './engine/adapter';
export { createEngine, PythonRdocxEngine, resolveArtifactPath } from './engine/adapter';
export {
  hasBlockingFailure,
  VERIFICATION_CHECK_IDS,
  verify,
  type VerificationCheckId,
  type VerificationInput,
} from './verifier';
export {
  assertDeclaredTypeMatches,
  assertParseUsable,
  collectLowConfidence,
  countTableCells,
  expandLogicalGrid,
  extensionOf,
  nodeIdFor,
  refineArtifactRef,
  toComplexContent,
  toComplexIR,
  toFormatProfile,
  toWarnings,
  type BudgetOverrun,
  type GridIssue,
  type LogicalCell,
  type LogicalGrid,
  type MappedContent,
  type MapperWarning,
} from './mapper';
export {
  isKnownWarningCode,
  parseEngineFailure,
  parseRawResult,
  SUPPORTED_PARSE_VERSION,
  type RawBlock,
  type RawContainerKind,
  type RawEngineFailure,
  type RawFormatKind,
  type RawIndicators,
  type RawLayout,
  type RawLayoutFragment,
  type RawLayoutPage,
  type RawPackagePart,
  type RawPackageProfile,
  type RawParagraphBlock,
  type RawParseResult,
  type RawRelationship,
  type RawTableBlock,
  type RawTableCell,
  type RawTableRow,
  type RawVMerge,
  type RawPageBreak,
  type RawSection,
  type RawFloat,
  type RawWarning,
} from './domain/docx-complex-parse';

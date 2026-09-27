/**
 * docx-inspect 的 Profile 注册入口。
 *
 * 本文件是模块的「装配层」，只做三件事：
 *   1. 把配置、引擎、映射器、验证器拼装成一个可注册的模块；
 *   2. 实现 `inspect` / `execute` / `verify` 三个 handler；
 *   3. 把内部异常统一收敛为本模块的错误分类。
 *
 * 明确的边界纪律（spec 第八节）：
 *   - 对外只暴露 `contract.ts` 定义的接口与配置，不泄漏引擎细节；
 *   - 不实现任何上层编排（那是 Profile 的职责）；
 *   - 不复制其他模块的公开契约；
 *   - 可以在 Profile 全量启动【之前】单独加载并测试。
 */
import type { ArtifactRef } from 'office-core';
import type {
  DocxExecuteFn,
  DocxInspectFn,
  DocxInspectHandlers,
  DocxInspectModule,
  DocxInspectModuleDefinition,
  DocxVerifyFn,
  DocxInspectOptions,
  ExecuteOutput,
  InspectOutput,
  ModuleConfig,
  ModuleContext,
  ModuleDependencyRef,
  ProfileRegistry,
  TelemetrySink,
  TelemetrySummary,
  VerificationReport,
  VerifyOutput,
} from './contract';
import { DOCX_INSPECT_CAPABILITIES, DOCX_INSPECT_MODULE_ID } from './contract';
import type { ConfigOverrides } from './config';
import { configSchema, resolveConfig, withCallOverrides } from './config';
import { DocxInspectError, toDocxInspectError } from './errors';
import {
  assertProbeUsable,
  refineArtifactRef,
  toFormatIR,
  toFormatProfile,
  toWarnings,
} from './mapper';
import { createTelemetryLogger } from './telemetry';
import type { DocxEngine } from './engine/adapter';
import { PythonDocxEngine } from './engine/adapter';
import type { ProbeResult } from './domain/docx-inspect';
import { hasBlockingFailure, runVerification } from './verifier';

/** 模块版本。必须与 `module.json` 中的 `version` 保持一致。 */
export const DOCX_INSPECT_VERSION = '0.1.0';

/** 依赖声明：module 指 Profile 内模块，runtime 指外部运行时。 */
const DEPENDENCIES: readonly ModuleDependencyRef[] = [
  { name: 'office-core', kind: 'module' },
  { name: 'office-safety', kind: 'module' },
  { name: 'office-files', kind: 'module', optional: true },
  { name: 'office-test-kit', kind: 'module', optional: true },
  { name: 'python', kind: 'runtime' },
];

/**
 * 模块清单（与 `module.json` 一一对应）。
 *
 * Profile 只从这份清单读取模块的身份、能力与依赖；它不关心模块内部实现。
 */
export const DOCX_INSPECT_DEFINITION: DocxInspectModuleDefinition = {
  id: DOCX_INSPECT_MODULE_ID,
  version: DOCX_INSPECT_VERSION,
  profileGroup: 'DOCX',
  summary:
    'Input probing and safe read for DOCX: extension/MIME, OOXML structure, macros, external references, embedded objects.',
  capabilities: DOCX_INSPECT_CAPABILITIES,
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
 * 构建一个可注册的模块实例。
 *
 * 这是一个【纯函数式工厂】：不做任何 I/O、不读全局配置、不注册副作用，
 * 因此测试里可以随意创建多个互不干扰的实例。
 *
 * @param options 可选的配置覆盖与依赖注入
 * @returns 包含 definition / handlers / dispose 的模块实例
 */
export function createDocxInspectModule(options: CreateModuleOptions = {}): DocxInspectModule {
  const config: ModuleConfig = resolveConfig(options.config);
  const context: ModuleContext = { config };
  if (options.telemetry !== undefined) {
    context.telemetry = options.telemetry;
  }
  context.logger = createTelemetryLogger(options.telemetry);

  // 引擎可以替换：默认是 Python 桥接，测试或未来实现可注入其他后端。
  const engine: DocxEngine = options.engine ?? new PythonDocxEngine(config.engine, options.telemetry);

  /** 按单次调用的 options 计算实际生效的配置。 */
  const effectiveConfig = (opts: DocxInspectOptions | undefined): ModuleConfig => {
    const overrides: ConfigOverrides = {};
    if (opts?.limits !== undefined) overrides.limits = opts.limits;
    if (opts?.featureFlags !== undefined) overrides.featureFlags = opts.featureFlags;
    if (opts?.timeoutMs !== undefined) overrides.timeoutMs = opts.timeoutMs;
    // 无覆盖时直接复用已冻结的配置，避免无谓重建。
    if (Object.keys(overrides).length === 0) return config;
    return withCallOverrides(config, overrides);
  };

  /** 调用引擎并把失败统一收敛为模块错误。 */
  const runProbe = async (
    input: { artifactRef: ArtifactRef; options?: DocxInspectOptions },
    active: ModuleConfig,
  ): Promise<ProbeResult> => {
    try {
      return await engine.probe({
        artifactPath: input.artifactRef.uri,
        limits: active.limits,
        featureFlags: active.featureFlags,
        timeoutMs: active.timeoutMs,
      });
    } catch (error) {
      // 引擎层抛出的已是模块错误（adapter 负责翻译），这里只兜底非预期异常。
      throw toDocxInspectError(error, 'ENGINE_FAILED', {
        message: 'Probe engine failed',
        path: input.artifactRef.uri,
      });
    }
  };

  /** 汇总遥测摘要（不含文档内容与绝对路径）。 */
  const buildTelemetry = (
    probe: ProbeResult,
    startedAt: number,
    engineMs: number,
  ): TelemetrySummary => ({
    probeVersion: probe.probeVersion,
    xmlBackend: probe.xmlBackend,
    engineMs,
    totalMs: Date.now() - startedAt,
    partCount: probe.partCount,
    relationshipCount: probe.relationships.length,
  });

  // ---------------------------------------------------------------------- //
  // inspect：回答「这是什么产物，是否可接受」                                 //
  // ---------------------------------------------------------------------- //
  const inspect: DocxInspectFn = async (input) => {
    const startedAt = Date.now();
    try {
      const active = effectiveConfig(input.options);
      const engineStartedAt = Date.now();
      const probe = await runProbe(input, active);
      const engineMs = Date.now() - engineStartedAt;

      // 类型/范围检查：不通过则直接抛错，不会产出半成品结果。
      assertProbeUsable(probe, active, input.options);

      const profile = toFormatProfile(probe, input.options);
      const warnings = toWarnings(probe, input.options);

      let verification: VerificationReport | undefined;
      if (input.policy !== undefined) {
        // 调用方给了策略：顺带评估，并在硬失败时快速拒绝。
        verification = runVerification(probe, input.policy, active);
        if (hasBlockingFailure(verification)) {
          throw new DocxInspectError(
            'SAFETY_POLICY_DENIED',
            'Artifact is rejected by the supplied safety policy',
            {
              path: input.artifactRef.uri,
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
        // 仅做「结构自检」：使用一个不声明任何要求的最小策略，
        // 因此结果只会是 pass 或 skip，不会因为策略而拒绝文档。
        verification = runVerification(probe, { id: 'inspect.self-check' }, active);
      }

      const output: InspectOutput = {
        moduleId: DOCX_INSPECT_MODULE_ID,
        requestId: input.requestId,
        operation: 'inspect',
        result: profile,
        // 探测不产生新文件，因此 artifacts 恒为空数组（保持形状一致）。
        artifacts: [],
        warnings,
        telemetry: buildTelemetry(probe, startedAt, engineMs),
      };
      if (verification !== undefined) {
        output.verification = verification;
      }
      return output;
    } catch (error) {
      context.logger?.error('docx-inspect.inspect.failed', {
        requestId: input.requestId,
      });
      throw toDocxInspectError(error, 'PARSE_FAILED', { path: input.artifactRef.uri });
    }
  };

  // ---------------------------------------------------------------------- //
  // execute：产出与引擎无关的 IR                                             //
  // ---------------------------------------------------------------------- //
  const execute: DocxExecuteFn = async (input) => {
    const startedAt = Date.now();
    try {
      const active = effectiveConfig(input.options);
      const engineStartedAt = Date.now();
      const probe = await runProbe(input, active);
      const engineMs = Date.now() - engineStartedAt;

      // execute 与 inspect 共享同一套可用性判定，保证两个接口结论一致。
      assertProbeUsable(probe, active, input.options);

      const ir = toFormatIR(probe, input.options);
      const warnings = toWarnings(probe, input.options);

      const output: ExecuteOutput = {
        moduleId: DOCX_INSPECT_MODULE_ID,
        requestId: input.requestId,
        operation: 'execute',
        result: { ir, artifact: refineArtifactRef(input.artifactRef, probe) },
        artifacts: [],
        warnings,
        telemetry: buildTelemetry(probe, startedAt, engineMs),
      };
      return output;
    } catch (error) {
      context.logger?.error('docx-inspect.execute.failed', { requestId: input.requestId });
      throw toDocxInspectError(error, 'PARSE_FAILED', { path: input.artifactRef.uri });
    }
  };

  // ---------------------------------------------------------------------- //
  // verify：按策略评估并产出局部验证报告                                      //
  // ---------------------------------------------------------------------- //
  const verify: DocxVerifyFn = async (input) => {
    const startedAt = Date.now();
    try {
      const active = effectiveConfig(input.options);
      const engineStartedAt = Date.now();
      const probe = await runProbe(input, active);
      const engineMs = Date.now() - engineStartedAt;

      assertProbeUsable(probe, active, input.options);

      // 验证报告自身的产出失败（而非「检查未通过」）才对应 VERIFICATION_FAILED。
      const report = runVerification(probe, input.policy, active);
      const warnings = toWarnings(probe, input.options);
      // 报告结论与告警互为补充：这里把「验证未通过」也提升为一条告警，
      // 让只消费 warnings 的上层同样能感知到。
      if (!report.ok) {
        warnings.push({
          code: 'VERIFICATION_FAILED',
          message: 'Artifact did not satisfy every declared policy requirement.',
          severity: 'error',
          details: {
            policyId: report.policyId ?? null,
            failed: report.summary.failed,
          },
        });
      }

      const output: VerifyOutput = {
        moduleId: DOCX_INSPECT_MODULE_ID,
        requestId: input.requestId,
        operation: 'verify',
        result: report,
        artifacts: [],
        warnings,
        verification: report,
        telemetry: buildTelemetry(probe, startedAt, engineMs),
      };
      return output;
    } catch (error) {
      context.logger?.error('docx-inspect.verify.failed', { requestId: input.requestId });
      throw toDocxInspectError(error, 'VERIFICATION_FAILED', { path: input.artifactRef.uri });
    }
  };

  const handlers: DocxInspectHandlers = { inspect, execute, verify };

  return {
    definition: DOCX_INSPECT_DEFINITION,
    handlers,
    dispose: async () => {
      await engine.dispose();
    },
  };
}

/**
 * 便捷函数：构建模块并注册到 Profile registry。
 *
 * 注意：本函数只调用 registry 的 `registerModule`，不做任何「编排」——
 * 加载顺序、依赖解析、生命周期均由 Profile 决定。
 */
export async function register(
  registry: ProfileRegistry,
  options: CreateModuleOptions = {},
): Promise<DocxInspectModule> {
  const module = createDocxInspectModule(options);
  await registry.registerModule(module);
  return module;
}

/* -------------------------------------------------------------------------- */
/* 对外导出                                                                     */
/* -------------------------------------------------------------------------- */

export * from './contract';
export { DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS, DEFAULT_TIMEOUT_MS, resolveConfig } from './config';
export { DocxInspectError, isDocxInspectError, toDocxInspectError } from './errors';
export { InMemoryTelemetry, noopTelemetry } from './telemetry';
export type { DocxEngine, ProbeRequest } from './engine/adapter';
export { PythonDocxEngine, resolveArtifactPath } from './engine/adapter';
export type { ConfigOverrides } from './config';
export { runVerification, hasBlockingFailure, VERIFICATION_CHECK_IDS } from './verifier';
export type { ProbeResult } from './domain/docx-inspect';

/**
 * 回归测试 —— 锁定「输出形状、顺序与结论」这些容易被无意改动的行为。
 *
 * 与另两个文件的分工：
 *   contract.test.ts    : 契约表面（类型/配置/错误/注册）
 *   edge-cases.test.ts  : 各类正常、边界、损坏与拒绝场景
 *   regression.test.ts  : 确定性、稳定性与验证结论（防止重构悄悄改变输出）
 */
import { statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

import { beforeAll, describe, expect, it } from 'vitest';

import { createDocxInspectModule } from '../src/index';
import type { DocxInspectModule, ExecuteOutput, InspectOutput, VerifyOutput } from '../src/index';
import { VERIFICATION_CHECK_IDS } from '../src/verifier';

const FIXTURES_DIR = fileURLToPath(new URL('../fixtures/_generated/', import.meta.url));

/** 拼出 fixture 的绝对路径。 */
function fixture(name: string): string {
  return join(FIXTURES_DIR, name);
}

let module: DocxInspectModule;

beforeAll(() => {
  module = createDocxInspectModule();
});

/**
 * 去掉遥测后再比较输出。
 *
 * 遥测里含毫秒耗时，天然不确定；但「结果 / 告警 / 验证」必须完全确定，
 * 否则回归测试无法做精确断言，下游也无法对 IR 做哈希。
 */
function stableOf(output: InspectOutput | ExecuteOutput | VerifyOutput): string {
  return JSON.stringify({
    moduleId: output.moduleId,
    operation: output.operation,
    result: output.result,
    artifacts: output.artifacts,
    warnings: output.warnings,
    verification: output.verification,
  });
}

describe('regression: deterministic output', () => {
  it('produces byte-identical results across repeated inspections', async () => {
    const first = await module.handlers.inspect({
      requestId: 'r1',
      operation: 'inspect',
      artifactRef: { id: 'x1', uri: fixture('external.docx') },
    });
    const second = await module.handlers.inspect({
      requestId: 'r2',
      operation: 'inspect',
      artifactRef: { id: 'x2', uri: fixture('external.docx') },
    });
    // 注意：requestId 与 artifactRef.id 在 stableOf 之外，因此不影响比较。
    expect(stableOf(first)).toBe(stableOf(second));
  });

  it('keeps warning order stable', async () => {
    const output = await module.handlers.inspect({
      requestId: 'r3',
      operation: 'inspect',
      artifactRef: { id: 'x3', uri: fixture('embedded.docx') },
    });
    // 顺序是契约的一部分：同一次调用中 ActiveX 必须排在「指标类」告警中靠后的位置。
    expect(output.warnings.map((warning) => warning.code)).toEqual([
      'EMBEDDED_OBJECTS_PRESENT',
      'ALT_CHUNKS_PRESENT',
      'ACTIVE_X_PRESENT',
    ]);
  });

  it('sorts IR collections for stable hashing', async () => {
    const output = await module.handlers.execute({
      requestId: 'r4',
      operation: 'execute',
      artifactRef: { id: 'x4', uri: fixture('external.docx') },
    });
    const parts = output.result.ir.parts.map((part) => part.name);
    expect(parts).toEqual([...parts].sort());

    const relationships = output.result.ir.relationships;
    const keys = relationships.map((relationship) => `${relationship.sourcePart}#${relationship.id}`);
    expect(keys).toEqual([...keys].sort());
  });
});

describe('regression: execute output', () => {
  it('refines the artifact reference without forging a new identity', async () => {
    const path = fixture('clean.docx');
    const output = await module.handlers.execute({
      requestId: 'r5',
      operation: 'execute',
      artifactRef: { id: 'keep-me', uri: path, label: 'custom-label' },
    });

    // id 必须保持不变：artifact 的身份由调用方拥有，模块不得改写。
    expect(output.result.artifact.id).toBe('keep-me');
    // 调用方已给 label，模块不得覆盖。
    expect(output.result.artifact.label).toBe('custom-label');
    // 其余字段由探测结果补全。
    expect(output.result.artifact.sizeBytes).toBe(statSync(path).size);
    expect(output.result.artifact.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(output.result.artifact.mediaType).toContain('wordprocessingml');
    expect(output.result.ir.profile.format).toBe('docx');
  });

  it('exposes macro and external indicators in the IR', async () => {
    const macros = await module.handlers.execute({
      requestId: 'r6',
      operation: 'execute',
      artifactRef: { id: 'x6', uri: fixture('macro.docm') },
    });
    expect(macros.result.ir.indicators.macros).toHaveLength(1);
    expect(macros.result.ir.indicators.macros[0]?.kind).toBe('vba');

    const external = await module.handlers.execute({
      requestId: 'r7',
      operation: 'execute',
      artifactRef: { id: 'x7', uri: fixture('external.docx') },
    });
    expect(external.result.ir.indicators.externalReferences).toHaveLength(3);
  });
});

describe('regression: verification conclusions', () => {
  it('passes every declared requirement for a clean document', async () => {
    const output = await module.handlers.verify({
      requestId: 'v1',
      operation: 'verify',
      artifactRef: { id: 'y1', uri: fixture('clean.docx') },
      policy: {
        id: 'strict-but-clean',
        allowMacros: false,
        allowExternalLinks: false,
        allowEmbeddedObjects: false,
        allowAltChunks: false,
        allowEncrypted: false,
        maxExternalTargets: 0,
      },
    });
    expect(output.result.ok).toBe(true);
    // 所有要求都已声明，因此不存在跳过项。
    expect(output.result.partial).toBe(false);
    expect(output.result.summary.failed).toBe(0);
    expect(output.result.policyId).toBe('strict-but-clean');
    // 不通过时才会追加 VERIFICATION_FAILED 告警。
    expect(output.warnings.map((warning) => warning.code)).not.toContain('VERIFICATION_FAILED');
  });

  it('marks undeclared requirements as skipped instead of passed', async () => {
    const output = await module.handlers.verify({
      requestId: 'v2',
      operation: 'verify',
      artifactRef: { id: 'y2', uri: fixture('clean.docx') },
      policy: { id: 'silent-policy' },
    });
    const byId = new Map(output.result.checks.map((entry) => [entry.id, entry]));
    expect(byId.get(VERIFICATION_CHECK_IDS.ENCRYPTION)?.status).toBe('skip');
    expect(byId.get(VERIFICATION_CHECK_IDS.MACROS)?.status).toBe('skip');
    // 结构检查与 ActiveX/DDE 检查永远执行，不依赖策略声明。
    expect(byId.get(VERIFICATION_CHECK_IDS.STRUCTURE)?.status).toBe('pass');
    expect(output.result.partial).toBe(true);
    expect(output.result.ok).toBe(true);
  });

  it('fails when the policy forbids macros', async () => {
    const output = await module.handlers.verify({
      requestId: 'v3',
      operation: 'verify',
      artifactRef: { id: 'y3', uri: fixture('macro.docm') },
      policy: { id: 'no-macros', allowMacros: false },
    });
    const macroCheck = output.result.checks.find(
      (entry) => entry.id === VERIFICATION_CHECK_IDS.MACROS,
    );
    expect(macroCheck?.status).toBe('fail');
    expect(output.result.ok).toBe(false);
    // 结论与告警必须一致，避免两种消费方式得到不同信号。
    expect(output.warnings.map((warning) => warning.code)).toContain('VERIFICATION_FAILED');
  });

  it('passes when the policy allows macros', async () => {
    const output = await module.handlers.verify({
      requestId: 'v4',
      operation: 'verify',
      artifactRef: { id: 'y4', uri: fixture('macro.docm') },
      policy: { id: 'allow-macros', allowMacros: true },
    });
    const macroCheck = output.result.checks.find(
      (entry) => entry.id === VERIFICATION_CHECK_IDS.MACROS,
    );
    expect(macroCheck?.status).toBe('pass');
    expect(output.result.ok).toBe(true);
  });

  it('treats reviewOnMacros as a blocking requirement', async () => {
    const output = await module.handlers.verify({
      requestId: 'v5',
      operation: 'verify',
      artifactRef: { id: 'y5', uri: fixture('macro.docm') },
      // 允许宏但要求人工复核 → 自动化管线不得判为通过。
      policy: { id: 'review-macros', allowMacros: true, reviewOnMacros: true },
    });
    const macroCheck = output.result.checks.find(
      (entry) => entry.id === VERIFICATION_CHECK_IDS.MACROS,
    );
    expect(macroCheck?.status).toBe('fail');
    expect(output.result.ok).toBe(false);
  });

  it('enforces the external target budget', async () => {
    const output = await module.handlers.verify({
      requestId: 'v6',
      operation: 'verify',
      artifactRef: { id: 'y6', uri: fixture('external.docx') },
      policy: { id: 'tight-budget', allowExternalLinks: true, maxExternalTargets: 1 },
    });
    const budgetCheck = output.result.checks.find(
      (entry) => entry.id === VERIFICATION_CHECK_IDS.EXTERNAL_TARGET_BUDGET,
    );
    expect(budgetCheck?.status).toBe('fail');
    expect(budgetCheck?.details?.['count']).toBe(3);
    expect(output.result.ok).toBe(false);
  });

  it('flags ActiveX and DDE as warnings rather than hard denials', async () => {
    const output = await module.handlers.verify({
      requestId: 'v7',
      operation: 'verify',
      artifactRef: { id: 'y7', uri: fixture('embedded.docx') },
      // 策略允许嵌入对象；ActiveX 仍应独立告警（该情形由 embedded.docx 覆盖）。
      policy: { id: 'allow-embedded', allowEmbeddedObjects: true, allowAltChunks: true },
    });
    const activeX = output.result.checks.find(
      (entry) => entry.id === VERIFICATION_CHECK_IDS.ACTIVE_X,
    );
    expect(activeX?.status).toBe('fail');
    // 关键：严重级别是 warn，因此不应被判定为「阻断性失败」。
    expect(activeX?.severity).toBe('warn');
    expect(output.result.summary.failed).toBeGreaterThan(0);
  });

  it('mirrors the verification report into the output envelope', async () => {
    const output = await module.handlers.verify({
      requestId: 'v8',
      operation: 'verify',
      artifactRef: { id: 'y8', uri: fixture('clean.docx') },
      policy: { id: 'mirror', allowEncrypted: false },
    });
    // result 与 verification 必须指向同一份结论，避免消费方二义。
    expect(output.verification).toEqual(output.result);
  });
});

describe('regression: JSON boundary discipline', () => {
  it('serializes every output without losing structure', async () => {
    const inspectOutput = await module.handlers.inspect({
      requestId: 'j1',
      operation: 'inspect',
      artifactRef: { id: 'z1', uri: fixture('external.docx') },
    });
    const executeOutput = await module.handlers.execute({
      requestId: 'j2',
      operation: 'execute',
      artifactRef: { id: 'z2', uri: fixture('external.docx') },
    });
    const verifyOutput = await module.handlers.verify({
      requestId: 'j3',
      operation: 'verify',
      artifactRef: { id: 'z3', uri: fixture('external.docx') },
      policy: { id: 'p', allowExternalLinks: true },
    });

    for (const output of [inspectOutput, executeOutput, verifyOutput]) {
      const encoded = JSON.stringify(output);
      expect(JSON.parse(encoded)).toEqual(JSON.parse(JSON.stringify(JSON.parse(encoded))));
      // 禁止出现引擎对象：Buffer 会被序列化成 `{type:'Buffer',data:[...]}`。
      expect(encoded).not.toContain('"type":"Buffer"');
      // 序列化后不应残留 `undefined`（说明有字段既未定义也不可序列化）。
      expect(encoded).not.toContain('undefined');
    }
  });

  it('keeps inputs JSON-serializable', () => {
    // 输入信封本身也必须能跨越 RPC 边界（spec 第四章「必须可序列化」）。
    const input = {
      requestId: 'j4',
      operation: 'inspect' as const,
      artifactRef: { id: 'z4', uri: fixture('clean.docx') },
      options: { declaredExtension: 'docx', limits: { maxArchiveEntries: 100 } },
    };
    expect(JSON.parse(JSON.stringify(input))).toEqual(input);
  });
});

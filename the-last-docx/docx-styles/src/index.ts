import type { ArtifactRef, Warning } from 'office-core';
import { createHash, timingSafeEqual } from 'node:crypto';
import type {
  DocxStylesModule,
  DocxStylesModuleDefinition,
  ExecuteInput,
  ExecuteOutput,
  InspectInput,
  InspectOutput,
  ModuleConfig,
  ModuleCreateOptions,
  ModuleDependencyRef,
  VerifyInput,
  VerifyOutput,
  VerificationCheck,
} from './contract';
import { DOCX_STYLES_CAPABILITIES, DOCX_STYLES_MODULE_ID } from './contract';
import { resolveConfig, withCallOverrides } from './config';
import { DocxStylesError, toDocxStylesError } from './errors';
import { Docx4jStylesEngine } from './engine/adapter';
import { makeVerificationReport } from './verifier';

export const DOCX_STYLES_VERSION = '0.1.0';

const DEPENDENCIES: readonly ModuleDependencyRef[] = [
  { name: 'office-core', kind: 'module' },
  { name: 'office-safety', kind: 'module' },
  { name: 'office-files', kind: 'module' },
  { name: 'office-test-kit', kind: 'module', optional: true },
  { name: '@docx4j/core-ts', kind: 'npm', version: '0.1.5' },
  { name: '@xmldom/xmldom', kind: 'npm', version: '0.9.12' },
];

export const DOCX_STYLES_CONFIG_SCHEMA: Record<string, unknown> = {
  type: 'object',
  additionalProperties: false,
  properties: {
    engine: { type: 'object', additionalProperties: false, properties: { driver: { type: 'string', enum: ['docx4j-core-ts'], default: 'docx4j-core-ts' } } },
    limits: {
      type: 'object', additionalProperties: false,
      properties: {
        maxInputBytes: { type: 'integer', minimum: 1, default: 67108864 },
        maxOutputBytes: { type: 'integer', minimum: 1, default: 134217728 },
        maxStyles: { type: 'integer', minimum: 1, default: 500 },
      },
    },
    timeoutMs: { type: 'integer', minimum: 1, default: 30000 },
    featureFlags: {
      type: 'object', additionalProperties: false,
      properties: {
        allowStyleDefinitions: { type: 'boolean', default: true },
        allowRedefineExisting: { type: 'boolean', default: true },
      },
    },
  },
};

export const DOCX_STYLES_DEFINITION: DocxStylesModuleDefinition = {
  id: DOCX_STYLES_MODULE_ID,
  version: DOCX_STYLES_VERSION,
  profileGroup: 'DOCX',
  summary: 'Named-style authoring and resolution over styles.xml, backed by @docx4j/core-ts.',
  capabilities: DOCX_STYLES_CAPABILITIES,
  dependencies: DEPENDENCIES,
  configSchema: DOCX_STYLES_CONFIG_SCHEMA,
};

function assertRequestId(requestId: string): void {
  if (typeof requestId !== 'string' || requestId.trim() === '') {
    throw new DocxStylesError('INVALID_INPUT', 'requestId must be a non-empty string.');
  }
}

function assertJson(value: unknown): void {
  try {
    const encoded = JSON.stringify(value);
    if (encoded === undefined) throw new Error('undefined');
    JSON.parse(encoded);
  } catch {
    throw new DocxStylesError('INVALID_INPUT', 'The request must be JSON serializable.');
  }
}

function checkSafety(
  options: ModuleCreateOptions,
  input: { artifactRef: ArtifactRef; policy?: InspectInput['policy']; operation: 'inspect' | 'execute' | 'verify' },
): Promise<void> {
  if (!input.policy) return Promise.resolve();
  if (!options.safetyGuard) {
    return Promise.reject(new DocxStylesError('SAFETY_POLICY_DENIED', 'A safety policy was supplied but no office-safety guard is connected.'));
  }
  return options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, policy: input.policy, operation: input.operation });
}

function reportFromEngine(result: { ok: boolean; checks: { id: string; ok: boolean; message: string }[] }) {
  const checks: VerificationCheck[] = result.checks.map((check) => ({
    id: check.id,
    status: check.ok ? 'pass' : 'fail',
    severity: check.ok ? 'info' : 'error',
    message: check.message,
  }));
  return makeVerificationReport(checks);
}

function outputBase<T>(requestId: string, operation: 'inspect' | 'execute' | 'verify', result: T, artifacts: ArtifactRef[] = [], warnings: Warning[] = []) {
  return { moduleId: DOCX_STYLES_MODULE_ID, requestId, operation, result, artifacts, warnings } as const;
}

async function withEngineTimeout<T>(task: () => Promise<T>, timeoutMs: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      task(),
      new Promise<T>((_resolve, reject) => {
        timer = setTimeout(() => reject(new DocxStylesError('ENGINE_TIMEOUT', 'The DOCX engine exceeded timeoutMs.')), timeoutMs);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

function assertArtifactDigest(ref: ArtifactRef, bytes: Uint8Array): void {
  if (!ref.sha256) return;
  const expected = ref.sha256.toLowerCase();
  const actual = createHash('sha256').update(bytes).digest('hex');
  if (!/^[a-f0-9]{64}$/.test(expected) || !timingSafeEqual(Buffer.from(expected), Buffer.from(actual))) {
    throw new DocxStylesError('FORMAT_MISMATCH', 'The artifact bytes do not match the artifactRef SHA-256.');
  }
}

/** Create an independently usable Profile module. Artifact I/O is delegated to office-files. */
export function createDocxStylesModule(options: ModuleCreateOptions): DocxStylesModule {
  const config: ModuleConfig = resolveConfig(options.config);
  const engine = new Docx4jStylesEngine();

  const activeConfig = (input: { options?: InspectInput['options'] }): ModuleConfig =>
    withCallOverrides(config, input.options);

  const inspect = async (input: InspectInput): Promise<InspectOutput> => {
    assertRequestId(input.requestId);
    assertJson(input);
    await checkSafety(options, { ...input, operation: 'inspect' });
    const active = activeConfig(input);
    const startedAt = Date.now();
    try {
      const bytes = await options.artifactStore.read(input.artifactRef, active.limits.maxInputBytes);
      const inventory = await withEngineTimeout(() => engine.inspect(bytes, active.limits), active.timeoutMs);
      options.telemetry?.record({
        name: 'docx-styles.inspect', timestampMs: startedAt, durationMs: Date.now() - startedAt,
        attributes: {
          engine: engine.name,
          styles: inventory.styles.length,
          unknownUndefined: inventory.references.unknownUndefined.length,
        },
      });
      return outputBase(input.requestId, 'inspect', inventory);
    } catch (error) {
      throw toDocxStylesError(error);
    }
  };

  const execute = async (input: ExecuteInput): Promise<ExecuteOutput> => {
    assertRequestId(input.requestId);
    assertJson(input);
    await checkSafety(options, { ...input, operation: 'execute' });
    const active = activeConfig(input);
    if (!input.plan || !Array.isArray(input.plan.styles)) {
      throw new DocxStylesError('INVALID_INPUT', 'A style plan must carry a styles array.');
    }
    const startedAt = Date.now();
    try {
      const source = await options.artifactStore.read(input.artifactRef, active.limits.maxInputBytes);
      assertArtifactDigest(input.artifactRef, source);
      const edited = await withEngineTimeout(
        () => engine.execute(source, input.plan.styles, active.limits, active.featureFlags),
        active.timeoutMs,
      );
      const rawVerification = await withEngineTimeout(
        () => engine.verify(edited.bytes, active.limits, { styles: input.plan.styles }),
        active.timeoutMs,
      );
      const verification = reportFromEngine(rawVerification);
      if (!verification.ok) {
        throw new DocxStylesError('VERIFICATION_FAILED', 'The saved document did not pass post-write verification.');
      }
      const artifact = await options.artifactStore.write({
        source: input.artifactRef,
        requestId: input.requestId,
        bytes: edited.bytes,
        suggestedName: `${input.artifactRef.label ?? input.artifactRef.id}.styled.docx`,
      });
      options.telemetry?.record({
        name: 'docx-styles.execute', timestampMs: startedAt, durationMs: Date.now() - startedAt,
        attributes: { engine: engine.name, definitions: edited.defined.length },
      });
      return {
        ...outputBase(input.requestId, 'execute', {
          artifact,
          defined: edited.defined,
          inventory: edited.inventory,
          verification,
        }, [artifact], []),
        verification,
      };
    } catch (error) {
      throw toDocxStylesError(error);
    }
  };

  const verify = async (input: VerifyInput): Promise<VerifyOutput> => {
    assertRequestId(input.requestId);
    assertJson(input);
    await checkSafety(options, { ...input, operation: 'verify' });
    const active = activeConfig(input);
    try {
      const bytes = await options.artifactStore.read(input.artifactRef, active.limits.maxInputBytes);
      assertArtifactDigest(input.artifactRef, bytes);
      const verification = reportFromEngine(
        await withEngineTimeout(() => engine.verify(bytes, active.limits, input.expectations), active.timeoutMs),
      );
      return { ...outputBase(input.requestId, 'verify', verification), verification };
    } catch (error) {
      throw toDocxStylesError(error);
    }
  };

  return {
    definition: DOCX_STYLES_DEFINITION,
    handlers: { inspect, execute, verify },
    async dispose() {},
  };
}

export * from './contract';
export * from './domain/docx-styles';
export { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS, DEFAULT_TIMEOUT_MS, resolveConfig } from './config';
export { DocxStylesError } from './errors';
export { Docx4jStylesEngine } from './engine/adapter';
export { PPR_CHILD_ORDER, RPR_CHILD_ORDER, STYLE_CHILD_ORDER, TBLPR_CHILD_ORDER, styleElement } from './engine/style-xml';

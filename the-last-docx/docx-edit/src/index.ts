import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import { createHash, timingSafeEqual } from 'node:crypto';
import type {
  DocxEditModule,
  DocxEditModuleDefinition,
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
import { DOCX_EDIT_CAPABILITIES, DOCX_EDIT_MODULE_ID } from './contract';
import { resolveConfig, withCallOverrides } from './config';
import { DocxEditError, toDocxEditError } from './errors';
import { Docx4jCoreTsEngine } from './engine/adapter';
import { makeVerificationReport } from './verifier';

export const DOCX_EDIT_VERSION = '0.2.0';

const DEPENDENCIES: readonly ModuleDependencyRef[] = [
  { name: 'office-core', kind: 'module' },
  { name: 'office-safety', kind: 'module' },
  { name: 'office-files', kind: 'module' },
  { name: 'office-test-kit', kind: 'module', optional: true },
  { name: 'docx-parse', kind: 'module' },
  { name: '@docx4j/core-ts', kind: 'npm', version: '0.1.5' },
];

export const DOCX_EDIT_CONFIG_SCHEMA: Record<string, unknown> = {
  type: 'object',
  additionalProperties: false,
  properties: {
    engine: { type: 'object', additionalProperties: false, properties: { driver: { type: 'string', enum: ['docx4j-core-ts'], default: 'docx4j-core-ts' } } },
    limits: {
      type: 'object', additionalProperties: false,
      properties: {
        maxInputBytes: { type: 'integer', minimum: 1, default: 67108864 },
        maxOutputBytes: { type: 'integer', minimum: 1, default: 134217728 },
        maxOperations: { type: 'integer', minimum: 1, default: 1000 },
      },
    },
    timeoutMs: { type: 'integer', minimum: 1, default: 30000 },
    featureFlags: {
      type: 'object', additionalProperties: false,
      properties: {
        allowTextReplacement: { type: 'boolean', default: true },
        allowComments: { type: 'boolean', default: true },
        allowTrackedChanges: { type: 'boolean', default: true },
        allowAcceptReject: { type: 'boolean', default: false },
        requireExpectedDigest: { type: 'boolean', default: true },
        allowFormatting: { type: 'boolean', default: true },
        allowTableFormatting: { type: 'boolean', default: true },
        allowBlockEdits: { type: 'boolean', default: true },
      },
    },
  },
};

export const DOCX_EDIT_DEFINITION: DocxEditModuleDefinition = {
  id: DOCX_EDIT_MODULE_ID,
  version: DOCX_EDIT_VERSION,
  profileGroup: 'DOCX',
  summary: 'Structured DOCX editing backed by @docx4j/core-ts, consuming docx-parse semantic anchors.',
  capabilities: DOCX_EDIT_CAPABILITIES,
  dependencies: DEPENDENCIES,
  configSchema: DOCX_EDIT_CONFIG_SCHEMA,
};

function assertRequestId(requestId: string): void {
  if (typeof requestId !== 'string' || requestId.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'requestId must be a non-empty string.');
  }
}

function assertJson(value: unknown): void {
  try {
    const encoded = JSON.stringify(value);
    if (encoded === undefined) throw new Error('undefined');
    JSON.parse(encoded);
  } catch {
    throw new DocxEditError('INVALID_INPUT', 'The request must be JSON serializable.');
  }
}

function checkSafety(
  options: ModuleCreateOptions,
  input: { artifactRef: ArtifactRef; policy?: InspectInput['policy']; operation: 'inspect' | 'execute' | 'verify' },
): Promise<void> {
  if (!input.policy) return Promise.resolve();
  if (!options.safetyGuard) {
    return Promise.reject(new DocxEditError('SAFETY_POLICY_DENIED', 'A safety policy was supplied but no office-safety guard is connected.'));
  }
  return options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, policy: input.policy, operation: input.operation });
}

function profileFromInspection(artifactRef: ArtifactRef, details: Awaited<ReturnType<Docx4jCoreTsEngine['inspect']>>): FormatProfile {
  return {
    format: 'docx',
    mediaType: artifactRef.mediaType ?? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    extension: 'docx',
    confidence: 1,
    container: 'zip',
    encrypted: false,
    signatures: ['zip:local-header', 'docx4j-core-ts:wordprocessingml'],
    features: {
      paragraphs: details.paragraphCount > 0,
      tables: details.tableCount > 0,
      comments: details.commentCount > 0,
      trackedChanges: details.trackedChangeCount > 0,
    },
    metadata: {
      paragraphCount: details.paragraphCount,
      tableCount: details.tableCount,
      commentCount: details.commentCount,
      trackedChangeCount: details.trackedChangeCount,
      byteLength: details.byteLength,
    },
  };
}

function reportFromEngine(result: Awaited<ReturnType<Docx4jCoreTsEngine['verify']>>) {
  const checks: VerificationCheck[] = result.checks.map((check) => ({
    id: check.id,
    status: check.ok ? 'pass' : 'fail',
    severity: check.ok ? 'info' : 'error',
    message: check.message,
  }));
  return makeVerificationReport(checks);
}

function outputBase<T>(requestId: string, operation: 'inspect' | 'execute' | 'verify', result: T, artifacts: ArtifactRef[] = [], warnings: Warning[] = []) {
  return { moduleId: DOCX_EDIT_MODULE_ID, requestId, operation, result, artifacts, warnings } as const;
}

async function withEngineTimeout<T>(task: () => Promise<T>, timeoutMs: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      task(),
      new Promise<T>((_resolve, reject) => {
        timer = setTimeout(() => reject(new DocxEditError('ENGINE_TIMEOUT', 'The DOCX engine exceeded timeoutMs.')), timeoutMs);
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
    throw new DocxEditError('STALE_TARGET', 'The artifact bytes do not match the artifactRef SHA-256.');
  }
}

/** Create an independently usable Profile module. Artifact I/O is delegated to office-files. */
export function createDocxEditModule(options: ModuleCreateOptions): DocxEditModule {
  const config: ModuleConfig = resolveConfig(options.config);
  const engine = new Docx4jCoreTsEngine();

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
      const inspection = await withEngineTimeout(() => engine.inspect(bytes, active.limits), active.timeoutMs);
      options.telemetry?.record({
        name: 'docx-edit.inspect', timestampMs: startedAt, durationMs: Date.now() - startedAt,
        attributes: { engine: engine.name, paragraphs: inspection.paragraphCount, tables: inspection.tableCount },
      });
      return outputBase(input.requestId, 'inspect', profileFromInspection(input.artifactRef, inspection));
    } catch (error) {
      throw toDocxEditError(error);
    }
  };

  const execute = async (input: ExecuteInput): Promise<ExecuteOutput> => {
    assertRequestId(input.requestId);
    assertJson(input);
    await checkSafety(options, { ...input, operation: 'execute' });
    const active = activeConfig(input);
    for (const edit of input.plan.edits) {
      if (active.featureFlags.requireExpectedDigest && !edit.target.anchor.digest) {
        throw new DocxEditError('INVALID_INPUT', 'The parser target must include its expected content digest.');
      }
      if (edit.kind === 'replaceText' && !active.featureFlags.allowTextReplacement) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Text replacement is disabled by Profile configuration.');
      }
      if (edit.kind === 'addComment' && !active.featureFlags.allowComments) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Comments are disabled by Profile configuration.');
      }
      if (edit.kind === 'replaceText' && edit.revision === 'track' && !active.featureFlags.allowTrackedChanges) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Tracked changes are disabled by Profile configuration.');
      }
      if (edit.kind === 'resolveRevisions' && !active.featureFlags.allowAcceptReject) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Accepting or rejecting revisions requires allowAcceptReject=true.');
      }
      if (edit.kind === 'formatParagraph' && !active.featureFlags.allowFormatting) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Paragraph formatting is disabled by Profile configuration.');
      }
      if (edit.kind === 'formatTable' && !active.featureFlags.allowTableFormatting) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Table formatting is disabled by Profile configuration.');
      }
      if (edit.kind === 'insertParagraph' && !active.featureFlags.allowBlockEdits) {
        throw new DocxEditError('UNSUPPORTED_OPERATION', 'Inserting paragraphs is disabled by Profile configuration.');
      }
    }
    const startedAt = Date.now();
    try {
      const source = await options.artifactStore.read(input.artifactRef, active.limits.maxInputBytes);
      assertArtifactDigest(input.artifactRef, source);
      const edited = await withEngineTimeout(
        () => engine.execute(source, input.plan, active.limits, active.featureFlags.allowTrackedChanges),
        active.timeoutMs,
      );
      const rawVerification = await withEngineTimeout(() => engine.verify(edited.bytes, active.limits, {
        paragraphTexts: edited.expectedParagraphTexts,
        commentTexts: edited.expectedCommentTexts,
        formats: edited.expectedFormats,
        tables: edited.expectedTableFormats,
        insertions: edited.expectedInsertions,
      }), active.timeoutMs);
      const verification = reportFromEngine(rawVerification);
      if (!verification.ok) {
        throw new DocxEditError('VERIFICATION_FAILED', 'The saved document did not pass post-edit verification.');
      }
      const artifact = await options.artifactStore.write({
        source: input.artifactRef,
        requestId: input.requestId,
        bytes: edited.bytes,
        suggestedName: `${input.artifactRef.label ?? input.artifactRef.id}.edited.docx`,
      });
      const warnings: Warning[] = [];
      options.telemetry?.record({
        name: 'docx-edit.execute', timestampMs: startedAt, durationMs: Date.now() - startedAt,
        attributes: { engine: engine.name, editCount: edited.edits.length, paragraphCount: edited.paragraphCount },
      });
      return {
        ...outputBase(input.requestId, 'execute', { artifact, edits: edited.edits, verification }, [artifact], warnings),
        verification,
      };
    } catch (error) {
      throw toDocxEditError(error);
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
      const verification = reportFromEngine(await withEngineTimeout(() => engine.verify(bytes, active.limits), active.timeoutMs));
      return { ...outputBase(input.requestId, 'verify', verification), verification };
    } catch (error) {
      throw toDocxEditError(error);
    }
  };

  return {
    definition: DOCX_EDIT_DEFINITION,
    handlers: { inspect, execute, verify },
    async dispose() {},
  };
}

export * from './contract';
export * from './domain/docx-edit';
export { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS, DEFAULT_TIMEOUT_MS, resolveConfig } from './config';
export { DocxEditError } from './errors';
export { Docx4jCoreTsEngine } from './engine/adapter';
export { targetFromDualIR, tableTargetFromDualIR } from './mapper';

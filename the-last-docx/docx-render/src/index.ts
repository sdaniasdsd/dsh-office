import type { ArtifactRef, FormatProfile, Warning } from 'office-core';
import type {
  DocxRenderModule, DocxRenderModuleDefinition, EngineRenderResult, ExecuteInput, InspectInput,
  ModuleConfig, ModuleCreateOptions, ModuleDependencyRef, ModuleOutput, RenderPage,
  RenderResult, VerifyInput,
} from './contract';
import { DOCX_RENDER_CAPABILITIES, DOCX_RENDER_MODULE_ID } from './contract';
import { resolveConfig, withCallOverrides } from './config';
import { DocxRenderError, toDocxRenderError } from './errors';
import { LibreOfficePopplerEngine, suggestedStem } from './engine/adapter';
import { findingsToWarnings, toFormatProfile } from './mapper';
import { recordRenderTelemetry } from './telemetry';
import { verifyRender } from './verifier';

export const DOCX_RENDER_VERSION = '0.1.0';
const DEPENDENCIES: readonly ModuleDependencyRef[] = [
  { name: 'office-core', kind: 'module' },
  { name: 'office-safety', kind: 'module' },
  { name: 'office-files', kind: 'module' },
  { name: 'docx-parse', kind: 'module' },
  { name: 'office-test-kit', kind: 'module', optional: true },
  { name: 'LibreOffice soffice', kind: 'runtime', optional: true },
  { name: 'Poppler pdftoppm', kind: 'runtime', optional: true },
];

export const DOCX_RENDER_CONFIG_SCHEMA: Record<string, unknown> = {
  type: 'object', additionalProperties: false,
  properties: {
    engine: { type: 'object', additionalProperties: false, properties: {
      driver: { type: 'string', enum: ['libreoffice-poppler'], default: 'libreoffice-poppler' },
      sofficePath: { type: 'string', default: 'soffice' }, pdftoppmPath: { type: 'string', default: 'pdftoppm' },
      tempRoot: { type: 'string' },
    } },
    limits: { type: 'object', additionalProperties: false, properties: {
      maxInputBytes: { type: 'integer', minimum: 1, default: 67108864 },
      maxPdfBytes: { type: 'integer', minimum: 1, default: 134217728 },
      maxTotalImageBytes: { type: 'integer', minimum: 1, default: 268435456 },
      maxPages: { type: 'integer', minimum: 1, default: 200 },
      dpi: { type: 'integer', minimum: 36, maximum: 300, default: 120 },
      thumbnailMaxEdge: { type: 'integer', minimum: 64, maximum: 2048, default: 480 },
    } },
    timeoutMs: { type: 'integer', minimum: 1, default: 120000 },
    featureFlags: { type: 'object', additionalProperties: false, properties: {
      emitPdf: { type: 'boolean', default: true }, emitThumbnails: { type: 'boolean', default: true },
      requireVisualReview: { type: 'boolean', default: true },
    } },
  },
};
export const DOCX_RENDER_DEFINITION: DocxRenderModuleDefinition = {
  id: DOCX_RENDER_MODULE_ID,
  version: DOCX_RENDER_VERSION,
  profileGroup: 'DOCX',
  summary: 'Renders DOCX to PDF and page images for agent visual review; does not parse or edit document structure.',
  capabilities: DOCX_RENDER_CAPABILITIES,
  dependencies: DEPENDENCIES,
  configSchema: DOCX_RENDER_CONFIG_SCHEMA,
};

function assertRequest(input: { requestId: string }): void {
  if (typeof input.requestId !== 'string' || input.requestId.trim() === '') {
    throw new DocxRenderError('INVALID_INPUT', 'requestId must be a non-empty string.');
  }
}
function assertJson(value: unknown): void {
  try {
    const encoded = JSON.stringify(value);
    if (encoded === undefined) throw new Error('undefined');
    JSON.parse(encoded);
  } catch {
    throw new DocxRenderError('INVALID_INPUT', 'The request must be JSON serializable.');
  }
}
function validateConfig(config: ModuleConfig): void {
  for (const [key, value] of Object.entries(config.limits)) {
    if (!Number.isInteger(value) || value < 1) throw new DocxRenderError('INVALID_INPUT', `${key} must be a positive integer.`);
  }
  if (config.limits.dpi < 36 || config.limits.dpi > 300) throw new DocxRenderError('INVALID_INPUT', 'dpi must be between 36 and 300.');
  if (config.limits.thumbnailMaxEdge < 64 || config.limits.thumbnailMaxEdge > 2048) throw new DocxRenderError('INVALID_INPUT', 'thumbnailMaxEdge must be between 64 and 2048.');
  if (!Number.isInteger(config.timeoutMs) || config.timeoutMs < 1) throw new DocxRenderError('INVALID_INPUT', 'timeoutMs must be a positive integer.');
}
async function checkSafety(options: ModuleCreateOptions, input: { artifactRef: ArtifactRef; policy?: InspectInput['policy']; operation: 'inspect' | 'execute' | 'verify' }): Promise<void> {
  if (!input.policy) return;
  if (!options.safetyGuard) throw new DocxRenderError('SAFETY_POLICY_DENIED', 'A safety policy was supplied but no office-safety guard is connected.');
  try {
    await options.safetyGuard.assertAllowed({ artifactRef: input.artifactRef, policy: input.policy, operation: input.operation });
  } catch (error) {
    throw new DocxRenderError('SAFETY_POLICY_DENIED', error instanceof Error ? error.message : 'The artifact was rejected by office-safety.');
  }
}
function outputBase<T>(operation: 'inspect' | 'execute' | 'verify', requestId: string, result: T, artifacts: ArtifactRef[] = [], warnings: Warning[] = []): ModuleOutput<T> {
  return { moduleId: DOCX_RENDER_MODULE_ID, requestId, operation, result, artifacts, warnings };
}
function validateFindings(findings: VerifyInput['visualFindings']): void {
  if (findings === undefined) return;
  if (!Array.isArray(findings)) throw new DocxRenderError('INVALID_INPUT', 'visualFindings must be an array.');
  for (const finding of findings) {
    if (!finding.id || !finding.message || !Number.isInteger(finding.page) || finding.page < 1) {
      throw new DocxRenderError('INVALID_INPUT', 'Each visual finding requires an id, message, and positive page number.');
    }
    if (finding.confidence !== undefined && (finding.confidence < 0 || finding.confidence > 1)) {
      throw new DocxRenderError('INVALID_INPUT', 'Visual finding confidence must be between 0 and 1.');
    }
  }
}
function validateSourceMap(input: VerifyInput): void {
  const findings = input.visualFindings ?? [];
  const hasSourceRefs = findings.some((finding) => (finding.sourceSemanticIds?.length ?? 0) > 0 || (finding.sourcePointers?.length ?? 0) > 0);
  if (!hasSourceRefs) return;
  const parseResult = input.parseResult;
  if (!parseResult) throw new DocxRenderError('SOURCE_MAP_REQUIRED', 'Findings with parser references require the corresponding docx-parse execute result.');
  const parsedArtifact = parseResult.artifact;
  if (parsedArtifact.id !== input.artifactRef.id || (parsedArtifact.sha256 && input.artifactRef.sha256 && parsedArtifact.sha256 !== input.artifactRef.sha256)) {
    throw new DocxRenderError('SOURCE_MAP_MISMATCH', 'The supplied parser result belongs to a different artifact revision.');
  }
  const sourceMap = parseResult.ir.content.sourceMap;
  for (const finding of findings) {
    for (const semanticId of finding.sourceSemanticIds ?? []) {
      if (!Object.hasOwn(sourceMap.bySemanticId, semanticId)) throw new DocxRenderError('SOURCE_MAP_MISMATCH', `Unknown parser semantic ID in finding ${finding.id}.`);
    }
    for (const pointer of finding.sourcePointers ?? []) {
      if (!Object.hasOwn(sourceMap.byPointer, pointer)) throw new DocxRenderError('SOURCE_MAP_MISMATCH', `Unknown parser source pointer in finding ${finding.id}.`);
    }
  }
}

/** Create an independently testable renderer. Filesystem access is isolated in the engine and ArtifactStore. */
export function createDocxRenderModule(options: ModuleCreateOptions): DocxRenderModule {
  if (!options?.artifactStore) throw new DocxRenderError('INVALID_INPUT', 'artifactStore is required.');
  const config = resolveConfig(options.config);
  validateConfig(config);
  const engine = options.engine ?? new LibreOfficePopplerEngine(config.engine);
  const activeConfig = (input: { options?: InspectInput['options'] }): ModuleConfig => {
    const active = withCallOverrides(config, input.options);
    validateConfig(active);
    return active;
  };
  const readArtifact = async (ref: ArtifactRef, maxBytes: number): Promise<Uint8Array> => {
    try {
      const bytes = await options.artifactStore.read(ref, maxBytes);
      if (!(bytes instanceof Uint8Array)) throw new Error('ArtifactStore.read must return Uint8Array.');
      if (bytes.byteLength > maxBytes) throw new DocxRenderError('LIMIT_EXCEEDED', 'The source artifact exceeds maxInputBytes.');
      return bytes;
    } catch (error) {
      if (error instanceof DocxRenderError) throw error;
      throw new DocxRenderError('ARTIFACT_NOT_FOUND', error instanceof Error ? error.message : 'Unable to read artifact.');
    }
  };
  const persistRender = async (ref: ArtifactRef, requestId: string, raw: EngineRenderResult, active: ModuleConfig): Promise<{ result: RenderResult; artifacts: ArtifactRef[] }> => {
    if (raw.pages.length === 0 || raw.pages.length > active.limits.maxPages) throw new DocxRenderError('ENGINE_PROTOCOL_ERROR', 'Renderer returned an invalid page count.');
    if (raw.pages.some((page, index) => page.pageNumber !== index + 1 || page.image.byteLength === 0 || (active.featureFlags.emitThumbnails && !page.thumbnail))) {
      throw new DocxRenderError('ENGINE_PROTOCOL_ERROR', 'Renderer returned missing, unordered, or empty page artifacts.');
    }
    const stem = suggestedStem(ref);
    const artifacts: ArtifactRef[] = [];
    let pdf: ArtifactRef | undefined;
    try {
      if (active.featureFlags.emitPdf) {
        pdf = await options.artifactStore.write({ source: ref, requestId, bytes: raw.pdf, suggestedName: `${stem}.preview.pdf` });
        artifacts.push(pdf);
      }
      const pages: RenderPage[] = [];
      for (const page of raw.pages) {
        const image = await options.artifactStore.write({ source: ref, requestId, bytes: page.image, suggestedName: `${stem}.page-${page.pageNumber}.png` });
        artifacts.push(image);
        let thumbnail: ArtifactRef | undefined;
        if (page.thumbnail && active.featureFlags.emitThumbnails) {
          thumbnail = await options.artifactStore.write({ source: ref, requestId, bytes: page.thumbnail, suggestedName: `${stem}.page-${page.pageNumber}.thumb.png` });
          artifacts.push(thumbnail);
        }
        pages.push({ pageNumber: page.pageNumber, image, ...(thumbnail ? { thumbnail } : {}) });
      }
      const result: RenderResult = {
        engine: engine.name,
        sourceArtifactId: ref.id,
        ...(ref.sha256 ? { sourceSha256: ref.sha256 } : {}),
        pageCount: pages.length,
        ...(pdf ? { pdf } : {}),
        pages,
        visualReview: 'pending',
        findings: [],
      };
      return { result, artifacts };
    } catch (error) {
      if (error instanceof DocxRenderError) throw error;
      throw new DocxRenderError('ARTIFACT_STORE_UNAVAILABLE', error instanceof Error ? error.message : 'Unable to persist render artifacts.');
    }
  };
  const render = async (input: ExecuteInput, active: ModuleConfig, startedAt: number, operation: 'execute' | 'verify' = 'execute') => {
    const bytes = await readArtifact(input.artifactRef, active.limits.maxInputBytes);
    await engine.inspect(bytes, input.artifactRef);
    const raw = await engine.render(bytes, active, input.artifactRef);
    if (raw.pdf.byteLength > active.limits.maxPdfBytes) throw new DocxRenderError('LIMIT_EXCEEDED', 'Rendered PDF exceeds maxPdfBytes.');
    const totalImageBytes = raw.pages.reduce((total, page) => total + page.image.byteLength + (page.thumbnail?.byteLength ?? 0), 0);
    if (totalImageBytes > active.limits.maxTotalImageBytes) throw new DocxRenderError('LIMIT_EXCEEDED', 'Rendered images exceed maxTotalImageBytes.');
    const saved = await persistRender(input.artifactRef, input.requestId, raw, active);
    recordRenderTelemetry(options.telemetry, { operation, requestId: input.requestId, engine: engine.name, startedAt, pageCount: saved.result.pageCount });
    return saved;
  };

  const inspect = async (input: InspectInput): Promise<ModuleOutput<FormatProfile>> => {
    assertRequest(input); assertJson(input);
    await checkSafety(options, { ...input, operation: 'inspect' });
    const active = activeConfig(input);
    const startedAt = Date.now();
    const bytes = await readArtifact(input.artifactRef, active.limits.maxInputBytes);
    try {
      const details = await engine.inspect(bytes, input.artifactRef);
      recordRenderTelemetry(options.telemetry, { operation: 'inspect', requestId: input.requestId, engine: engine.name, startedAt });
      return outputBase('inspect', input.requestId, toFormatProfile(input.artifactRef, details));
    } catch (error) { throw toDocxRenderError(error); }
  };
  const execute = async (input: ExecuteInput) => {
    assertRequest(input); assertJson(input);
    await checkSafety(options, { ...input, operation: 'execute' });
    const active = activeConfig(input);
    const startedAt = Date.now();
    try {
      const saved = await render(input, active, startedAt);
      const warnings: Warning[] = [{ code: 'FONT_SUBSTITUTION_POSSIBLE', severity: 'warn', message: 'Page layout reflects the configured renderer and available fonts; it may differ from Microsoft Word.' }];
      return outputBase('execute', input.requestId, saved.result, saved.artifacts, warnings);
    } catch (error) { throw toDocxRenderError(error); }
  };
  const verify = async (input: VerifyInput) => {
    assertRequest(input); assertJson(input); validateFindings(input.visualFindings);
    validateSourceMap(input);
    await checkSafety(options, { ...input, operation: 'verify' });
    const active = activeConfig(input);
    const startedAt = Date.now();
    try {
      let renderResult = input.renderResult;
      let artifacts: ArtifactRef[] = [];
      if (renderResult && (renderResult.sourceArtifactId !== input.artifactRef.id || (renderResult.sourceSha256 && input.artifactRef.sha256 && renderResult.sourceSha256 !== input.artifactRef.sha256))) {
        throw new DocxRenderError('STALE_RENDER_RESULT', 'The supplied render result belongs to a different artifact revision.');
      }
      if (!renderResult) {
        const rendered = await render({ ...input, operation: 'execute' }, active, startedAt, 'verify');
        renderResult = rendered.result;
        artifacts = rendered.artifacts;
      }
      for (const finding of input.visualFindings ?? []) {
        if (finding.page > renderResult.pageCount) throw new DocxRenderError('INVALID_INPUT', `Visual finding ${finding.id} refers to a page outside the rendered document.`);
      }
      const verification = verifyRender({
        pageCount: renderResult.pageCount,
        renderedPages: renderResult.pages.map((page) => page.pageNumber),
        requireVisualReview: active.featureFlags.requireVisualReview,
        findings: input.visualFindings,
      });
      const warnings: Warning[] = [];
      if (verification.partial) warnings.push({ code: 'VISUAL_REVIEW_PENDING', severity: 'warn', message: 'Rendering succeeded; an agent or user must inspect the page images before visual verification is complete.' });
      warnings.push(...findingsToWarnings(input.visualFindings ?? []));
      const result = { ...verification };
      const output = outputBase('verify', input.requestId, result, artifacts, warnings);
      if (input.renderResult) recordRenderTelemetry(options.telemetry, { operation: 'verify', requestId: input.requestId, engine: engine.name, startedAt, pageCount: renderResult.pageCount });
      return { ...output, verification };
    } catch (error) { throw toDocxRenderError(error); }
  };
  return {
    definition: DOCX_RENDER_DEFINITION,
    handlers: { inspect, execute, verify },
    dispose: async () => { await engine.dispose(); },
  };
}

export interface ProfileRegistryLike { registerModule(module: DocxRenderModule): void | Promise<void> }
export async function register(registry: ProfileRegistryLike, options: ModuleCreateOptions): Promise<DocxRenderModule> {
  const module = createDocxRenderModule(options);
  await registry.registerModule(module);
  return module;
}

export * from './contract';
export { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS, DEFAULT_TIMEOUT_MS, resolveConfig } from './config';
export { DocxRenderError } from './errors';
export { LibreOfficePopplerEngine } from './engine/adapter';
export { makeVerificationReport, verifyRender } from './verifier';

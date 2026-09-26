import { z } from 'zod';
import { artifactSchema, sourceSchema, renderResultSchema, makeReport, PdfModuleError, assertSameSource } from '@dsh-office-profile/pdf-contracts';
import type { ArtifactRef, EngineContext, ModuleOptions, VerificationCheck } from '@dsh-office-profile/pdf-contracts';
import { deliveryPlanSchema, plan } from './plans';
import { inspectPdf } from './native';
import { readArtifact } from './process';

const manifestSchema = z.object({ schema: z.literal('pdf-delivery/v1'), documentId: z.string(), revision: z.number().int().positive(),
  document: artifactSchema, parentManifest: artifactSchema.optional(), preview: renderResultSchema.optional(),
  evidence: z.array(artifactSchema), review: z.enum(['pending','reviewed','failed']),
  visualReview: deliveryPlanSchema.shape.visualReview }).strict();
async function read(ref: ArtifactRef, context: EngineContext, options: ModuleOptions) {
  if (!options.files || !ref.sha256 || ref.sizeBytes === undefined) throw new PdfModuleError('INVALID_INPUT', 'Delivery references require reader, hash and size.');
  return readArtifact(ref, context, options, context.config.limits.maxOutputBytes);
}
function json(bytes: Uint8Array) {
  try { return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)); }
  catch { throw new PdfModuleError('ENGINE_PROTOCOL_ERROR', 'Invalid delivery JSON.'); }
}
export async function deliver(context: EngineContext, options: ModuleOptions) {
  if (!options.writer) throw new PdfModuleError('INVALID_INPUT', 'Manifest writer required.');
  const input = plan(deliveryPlanSchema, context.request.payload);
  const profile = await inspectPdf(context);
  const source = context.source!;
  const document = { ...source.ref, sha256: source.identity.sha256, sizeBytes: source.bytes.length, mediaType: 'application/pdf' };
  if (input.revision === 1 && input.parentManifest) throw new PdfModuleError('INVALID_INPUT', 'Initial revision cannot have parent.');
  if (input.revision > 1) {
    if (!input.parentManifest) throw new PdfModuleError('INVALID_INPUT', 'Previous manifest required.');
    const previous = manifestSchema.parse(json(await read(input.parentManifest, context, options)));
    if (previous.documentId !== input.documentId || previous.revision !== input.revision - 1)
      throw new PdfModuleError('ARTIFACT_CONFLICT', 'Parent belongs to a different document/revision.');
  }
  let total = source.bytes.length, refs = 0;
  const boundedRead = async (ref: ArtifactRef) => {
    if (++refs > context.config.limits.maxArtifacts) throw new PdfModuleError('LIMIT_EXCEEDED', 'Delivery reference count exceeded.');
    total += ref.sizeBytes ?? 0;
    if (total > context.config.limits.maxTotalOutputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Delivery byte budget exceeded.');
    return read(ref, context, options);
  };
  if (input.preview) {
    assertSameSource(input.preview.source, source.identity);
    if (input.preview.pageCount !== profile.pageCount || input.preview.pages.length !== profile.pageCount || input.preview.pages.some((p, i) => p.pageNumber !== i + 1))
      throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Preview page set does not cover the source.');
    for (const page of input.preview.pages) await boundedRead(page.image);
  }
  for (const ref of input.evidence) {
    const evidence = z.object({ schema: z.literal('pdf-evidence/v1'), source: sourceSchema, payload: z.unknown() }).strict().parse(json(await boundedRead(ref)));
    assertSameSource(evidence.source, source.identity);
  }
  let review: 'pending' | 'reviewed' | 'failed' = 'pending';
  if (input.visualReview) {
    assertSameSource(input.visualReview.source, source.identity);
    if (!input.preview) throw new PdfModuleError('INVALID_INPUT', 'Visual review requires a source-bound preview.');
    const pages = new Set(input.visualReview.reviewedPages);
    if ([...pages].some(n => n > profile.pageCount!) || input.visualReview.findings.some(f => f.page > profile.pageCount!))
      throw new PdfModuleError('INVALID_INPUT', 'Visual review contains nonexistent pages.');
    review = input.visualReview.findings.some(f => f.severity === 'error') ? 'failed' : pages.size === profile.pageCount ? 'reviewed' : 'pending';
  }
  const manifest = manifestSchema.parse({ schema: 'pdf-delivery/v1', ...input, document, review });
  const bytes = Buffer.from(JSON.stringify(manifest));
  const manifestRef = await options.writer.commitManifest({ documentId: input.documentId, revision: input.revision,
    ...(input.parentManifest ? { parentSha256: input.parentManifest.sha256! } : {}), requestId: context.request.requestId, bytes, signal: context.signal });
  return { result: { manifestRef, document: source.identity, revision: input.revision, review }, artifacts: [manifestRef],
    warnings: review === 'pending' ? [{ code: 'VISUAL_REVIEW_PENDING', severity: 'info', message: 'Delivery is complete; visual review is still pending.' }] : [] };
}
export async function verifyDelivery(context: EngineContext, options: ModuleOptions) {
  const checks: VerificationCheck[] = [];
  let total = 0, count = 0;
  const visited = new Set<string>();
  const boundedRead = async (ref: ArtifactRef) => {
    if (++count > context.config.limits.maxArtifacts) throw new PdfModuleError('LIMIT_EXCEEDED', 'Version-chain reference budget exceeded.');
    total += ref.sizeBytes ?? 0;
    if (total > context.config.limits.maxTotalOutputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Version-chain byte budget exceeded.');
    return read(ref, context, options);
  };
  let current = manifestSchema.parse(json(context.source!.bytes));
  const latest = current;
  while (true) {
    const key = `${current.documentId}:${current.revision}`;
    if (visited.has(key)) throw new PdfModuleError('ARTIFACT_CONFLICT', 'Manifest version cycle.');
    visited.add(key);
    const bytes = await boundedRead(current.document);
    const identity = { id: current.document.id, sha256: current.document.sha256! };
    const profile = await inspectPdf({ ...context, source: { ref: current.document, bytes, identity } });
    if (current.preview) {
      assertSameSource(current.preview.source, identity);
      if (current.preview.pageCount !== profile.pageCount || current.preview.pages.length !== profile.pageCount || current.preview.pages.some((p,i) => p.pageNumber !== i+1))
        throw new PdfModuleError('SOURCE_MAP_MISMATCH', 'Manifest preview has an invalid page set.');
      for (const page of current.preview.pages) await boundedRead(page.image);
    }
    for (const ref of current.evidence) {
      const evidence = z.object({ schema: z.literal('pdf-evidence/v1'), source: sourceSchema, payload: z.unknown() }).strict().parse(json(await boundedRead(ref)));
      assertSameSource(evidence.source, identity);
    }
    if (current.visualReview) assertSameSource(current.visualReview.source, identity);
    const review = current.visualReview;
    const expectedState = !review ? 'pending' : review.findings.some(f => f.severity === 'error') ? 'failed'
      : new Set(review.reviewedPages).size === profile.pageCount ? 'reviewed' : 'pending';
    if (current.review !== expectedState || (review && (!current.preview || review.reviewedPages.some(n => n > profile.pageCount!) || review.findings.some(f => f.page > profile.pageCount!))))
      throw new PdfModuleError('VERIFICATION_FAILED', 'Manifest review claim is unsupported.');
    if (current.revision === 1) {
      if (current.parentManifest) throw new PdfModuleError('ARTIFACT_CONFLICT', 'Initial manifest has parent.');
      break;
    }
    if (!current.parentManifest) throw new PdfModuleError('ARTIFACT_CONFLICT', 'Manifest chain is incomplete.');
    const previous = manifestSchema.parse(json(await boundedRead(current.parentManifest)));
    if (previous.documentId !== current.documentId || previous.revision !== current.revision - 1)
      throw new PdfModuleError('ARTIFACT_CONFLICT', 'Manifest chain mismatch.');
    current = previous;
  }
  checks.push({ id: 'delivery.references', severity: 'error', status: 'pass', message: 'All artifact bytes and evidence source bindings verified.' },
    { id: 'delivery.chain', severity: 'error', status: 'pass', message: `${visited.size} consecutive revisions verified.` },
    { id: 'visual.review', severity: 'error', status: latest.review === 'reviewed' ? 'pass' : latest.review === 'failed' ? 'fail' : 'skip', message: 'Source-bound caller review status verified.' });
  return { result: makeReport(checks), artifacts: [], warnings: [] };
}

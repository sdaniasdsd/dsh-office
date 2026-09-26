import { createHash } from 'node:crypto';
import { PdfModuleError } from './errors';
import type { ErrorCode } from './errors';
import type { VerificationCheck, VerificationReport, PdfFormatProfile, PdfSafetyPolicy } from './contract';

export function sha256(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex');
}
/** Reject values JSON.stringify silently changes (NaN, undefined, Date, bytes, holes). */
export function assertJson(value: unknown, code: ErrorCode = 'INVALID_INPUT'): void {
  const ancestors = new Set<object>();
  const visit = (item: unknown, depth: number): void => {
    if (depth > 100) throw new PdfModuleError(code, 'JSON nesting exceeds 100.');
    if (item === null || typeof item === 'string' || typeof item === 'boolean') return;
    if (typeof item === 'number' && Number.isFinite(item)) return;
    if (typeof item !== 'object') throw new PdfModuleError(code, 'Expected lossless JSON values.');
    if (ancestors.has(item)) throw new PdfModuleError(code, 'Circular JSON value.');
    const proto = Object.getPrototypeOf(item);
    if (!Array.isArray(item) && proto !== Object.prototype && proto !== null)
      throw new PdfModuleError(code, 'Class instances and binary data cannot cross the JSON boundary.');
    if (Object.getOwnPropertySymbols(item).length) throw new PdfModuleError(code, 'Symbol keys are not JSON.');
    ancestors.add(item);
    if (Array.isArray(item)) {
      if (Object.keys(item).length !== item.length) throw new PdfModuleError(code, 'Sparse or decorated array.');
      for (let i = 0; i < item.length; i++) visit(item[i], depth + 1);
    } else {
      for (const descriptor of Object.values(Object.getOwnPropertyDescriptors(item))) {
        if (!descriptor.enumerable || !('value' in descriptor)) throw new PdfModuleError(code, 'Only enumerable data properties are allowed.');
        visit(descriptor.value, depth + 1);
      }
    }
    ancestors.delete(item);
  };
  visit(value, 0);
}
export function makeReport(checks: VerificationCheck[]): VerificationReport {
  return {
    ok: !checks.some(c => c.status === 'fail' && c.severity === 'error'),
    partial: checks.some(c => c.status === 'skip'), checks,
    summary: { total: checks.length, passed: checks.filter(c => c.status === 'pass').length,
      failed: checks.filter(c => c.status === 'fail').length, skipped: checks.filter(c => c.status === 'skip').length },
  };
}
export function verifyPolicy(profile: PdfFormatProfile, policy?: PdfSafetyPolicy): VerificationReport {
  const checks: VerificationCheck[] = [];
  for (const [id, allow, present] of [
    ['encrypted', policy?.allowEncrypted, profile.encrypted],
    ['javascript', policy?.allowJavaScript, profile.features.javascript],
    ['embeddedFiles', policy?.allowEmbeddedFiles, profile.features.embeddedFiles],
    ['externalLinks', policy?.allowExternalLinks, profile.features.externalLinks],
  ] as const) {
    checks.push({ id: `policy.${id}`, severity: 'error',
      status: allow === undefined ? 'skip' : allow || present === false ? 'pass' : 'fail',
      message: allow === undefined ? 'Requirement not declared.' : present === null && !allow
        ? 'Required absence could not be established.' : 'Evaluated against observed PDF features.' });
  }
  checks.push({ id: 'policy.maxPages', severity: 'error',
    status: policy?.maxPages === undefined ? 'skip' : profile.pageCount !== null && profile.pageCount <= policy.maxPages ? 'pass' : 'fail',
    message: 'Unknown page count does not satisfy a declared page limit.' });
  return makeReport(checks);
}

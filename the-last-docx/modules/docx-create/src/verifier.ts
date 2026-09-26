import type { ModuleConfig, VerificationReport } from './contract';
import type { SafetyPolicy } from 'office-safety';
import { checkPackage } from './engine/package-check';
import { normalizeError } from './errors';
export function verifyPackage(bytes: Uint8Array, config: ModuleConfig, policy?: SafetyPolicy): VerificationReport {
  const checks: VerificationReport['checks'] = [];
  try {
    const facts = checkPackage(bytes, config, policy);
    checks.push({ id: 'package-structure', status: 'pass', severity: 'info', message: 'DOCX parts, internal relationship targets and paragraph IDs checked; not full OOXML schema validation.' });
    checks.push({ id: 'fields', status: facts.fields ? 'skip' : 'pass', severity: 'info', message: facts.fields ? 'Field values and TOC require a layout application to update.' : 'No fields to update.' });
  } catch (error) {
    const e = normalizeError(error);
    checks.push({ id: 'package-structure', status: 'fail', severity: 'error', message: `${e.code}: ${e.message}` });
  }
  checks.push({ id: 'visual-layout', status: 'skip', severity: 'info', message: 'Send the output artifact to docx-render for page images and visual review.' });
  const passed = checks.filter(c => c.status === 'pass').length;
  const failed = checks.filter(c => c.status === 'fail').length;
  const skipped = checks.filter(c => c.status === 'skip').length;
  return { ok: failed === 0, partial: skipped > 0, checks, summary: { total: checks.length, passed, failed, skipped } };
}

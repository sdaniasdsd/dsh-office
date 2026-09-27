import type { VerificationCheck, VerificationReport, VisualFinding } from './contract';

export function makeVerificationReport(checks: VerificationCheck[]): VerificationReport {
  const passed = checks.filter((check) => check.status === 'pass').length;
  const failed = checks.filter((check) => check.status === 'fail').length;
  const skipped = checks.filter((check) => check.status === 'skip').length;
  return { ok: failed === 0, partial: skipped > 0, checks, summary: { total: checks.length, passed, failed, skipped } };
}
export function verifyRender(input: {
  pageCount: number;
  renderedPages: number[];
  requireVisualReview: boolean;
  findings?: VisualFinding[];
}): VerificationReport {
  const orderedPages = [...input.renderedPages].sort((left, right) => left - right);
  const complete = input.pageCount > 0
    && orderedPages.length === input.pageCount
    && orderedPages.every((page, index) => page === index + 1);
  const checks: VerificationCheck[] = [{
    id: 'render.pages.complete', status: complete ? 'pass' : 'fail', severity: 'error',
    message: complete ? 'Every page has a rendered image.' : 'The rendered image set is incomplete.',
    details: { expectedPages: input.pageCount, renderedPages: input.renderedPages.length },
  }];
  if (input.requireVisualReview && input.findings === undefined) {
    checks.push({ id: 'visual.review', status: 'skip', severity: 'warn', message: 'Page images still need visual review; no findings were supplied.' });
  } else {
    const findings = input.findings ?? [];
    const failures = findings.filter((finding) => finding.severity === 'error').length;
    const warnings = findings.filter((finding) => finding.severity === 'warn').length;
    checks.push({
      id: 'visual.findings', status: failures > 0 ? 'fail' : 'pass', severity: failures > 0 ? 'error' : warnings > 0 ? 'warn' : 'info',
      message: findings.length === 0 ? 'No visual defects were reported by the reviewer.' : `${findings.length} visual finding(s) were reported.`,
      details: { findings: findings.length, errors: failures, warnings },
    });
  }
  return makeVerificationReport(checks);
}

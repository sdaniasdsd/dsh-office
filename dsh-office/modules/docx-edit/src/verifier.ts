import type { VerificationCheck, VerificationReport } from './contract';

export function makeVerificationReport(checks: VerificationCheck[]): VerificationReport {
  const passed = checks.filter((check) => check.status === 'pass').length;
  const failed = checks.filter((check) => check.status === 'fail').length;
  const skipped = checks.filter((check) => check.status === 'skip').length;
  return {
    ok: failed === 0,
    partial: skipped > 0,
    checks,
    summary: { total: checks.length, passed, failed, skipped },
  };
}

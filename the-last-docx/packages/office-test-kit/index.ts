// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.

  /** `VerificationReport` 必须满足的形状契约。 */
  export interface VerificationReportLike {
    /** 是否所有「已声明」的检查都通过。 */
    ok: boolean;
    /** 是否为部分验证（存在跳过项或无法完成的检查）。 */
    partial: boolean;
    checks: unknown[];
    summary: {
      total: number;
      passed: number;
      failed: number;
      skipped: number;
    };
  }

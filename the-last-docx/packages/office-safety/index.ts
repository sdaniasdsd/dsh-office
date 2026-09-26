// Promoted from docx-inspect/types/office-deps.d.ts; single shared owner.

  /**
   * `verify` 接口所评估的安全策略。
   *
   * 所有开关都是可选的：缺省表示「本策略未声明该要求」，对应的检查会被标记为
   * `skip`（跳过）而不是 `pass`（通过）。这样策略作者能明确区分
   * 「我确认允许」与「我根本没表态」。
   */
  export interface SafetyPolicy {
    maxBlocks?: number;
    maxTableCells?: number;
    /** 策略标识，会回填到验证报告中。 */
    id: string;
    /** 是否允许文档携带宏。 */
    allowMacros?: boolean;
    /** 是否允许外部链接（远程模板、超链接等）。 */
    allowExternalLinks?: boolean;
    /** 是否允许嵌入对象。 */
    allowEmbeddedObjects?: boolean;
    /** 是否允许 altChunk（可引入外部内容块）。 */
    allowAltChunks?: boolean;
    /** 是否允许加密产物。 */
    allowEncrypted?: boolean;
    /** 外部关系目标数量的上限。 */
    maxExternalTargets?: number;
    /** 为 true 时，携带宏的文档必须转人工复核。 */
    reviewOnMacros?: boolean;
  }

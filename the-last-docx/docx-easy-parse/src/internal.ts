/**
 * 内部接缝（internal seams）——【不是】公开契约的一部分。
 *
 * 为什么单独开一个文件，而不是继续从 `index.ts` 里导出：
 *
 * `package.json` 的 `exports` 只映射了 `"."` → `./src/index.ts`，因此对模块外的
 * 调用方而言，`index.ts` 就是全部可见面。把映射、领域派生、验证、遥测与具体引擎
 * 实现挂在 `index.ts` 上，会让它们事实上变成公开 API：任何一次内部重命名或签名调整
 * 都会变成破坏性变更。
 *
 * 这里的划分标准是「调用方是否需要」：
 *   - 公开（index.ts）：加载、配置、调用、扩展四件事所需的东西；
 *   - 内部（本文件）：实现接缝，可自由重构。
 *
 * 仓库内的测试与本模块自身的其他文件通过相对路径引用本文件，从而既能复用这些
 * 接缝，又不把它们暴露给外部。
 */

/* --- 配置：默认值与调用级覆盖 ---------------------------------------------- */
// `resolveConfig` / `configSchema` 是公开的；默认值和覆盖合并是实现细节。

export {
  DEFAULT_DEEP_TIMEOUT_MS,
  DEFAULT_ENGINE,
  DEFAULT_FEATURE_FLAGS,
  DEFAULT_LIMITS,
  DEFAULT_TIMEOUT_MS,
  withCallOverrides,
} from './config';

/* --- 引擎：具体实现 -------------------------------------------------------- */
// 端口（DocxEngine / ParseRequest）是公开的扩展点，实现在这里。
// 换成别的后端时，只有这个文件与 adapter.ts 变动，公开契约不动。

export {
  DoclingEngine,
  PythonDocxEngine,
  createDeepEngine,
  defaultDoclingScriptPath,
  defaultScriptPath,
  resolveArtifactPath,
} from './engine/adapter';

/* --- 映射：裁决 / 派生 / 归一化 -------------------------------------------- */

export {
  ISSUE_TO_WARNING,
  assertParseUsable,
  evaluateDeclaredHints,
  refineArtifactRef,
  toDocxParseIR,
  toFormatProfile,
  toWarnings,
} from './mapper';
export type { HintConflict, HintEvaluation } from './mapper';

/* --- 验证 ----------------------------------------------------------------- */

export { VERIFICATION_CHECK_IDS, hasBlockingFailure, runVerification } from './verifier';

/* --- 遥测 ----------------------------------------------------------------- */

export { InMemoryTelemetry, createTelemetryLogger, noopTelemetry, withSpan } from './telemetry';

/* --- 领域：协议校验与派生纯函数 -------------------------------------------- */
// `parseParseResult` 是引擎载荷的信任边界，`ParseProtocolError` 是它的失败信号；
// 两者都对上层隐藏——调用方只会看到翻译后的 `DocxParseError`。

export {
  ParseProtocolError,
  buildOutline,
  computeCounts,
  hasVisibleContent,
  isHeading,
  parseParseResult,
} from './domain/docx-parse';

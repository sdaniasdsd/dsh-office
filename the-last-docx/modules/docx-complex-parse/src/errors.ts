/**
 * 本模块的错误分类体系。
 *
 * 关键纪律：底层引擎只能上报一个「问题」，而「问题 → 错误码」的映射由 mapper
 * 决定，引擎无权选择错误码。这样更换引擎实现时，上层看到的错误语义完全一致。
 */
import { DOCX_COMPLEX_PARSE_ERROR_CODES, type DocxComplexParseErrorCode } from './contract';

/** 构造参数。 */
export interface DocxComplexParseErrorOptions {
  /** 结构化上下文；必须可 JSON 序列化，且不得包含文档正文。 */
  details?: Record<string, unknown>;
  /** 上游原因；不会被序列化。 */
  cause?: unknown;
}

/** 本模块唯一的错误类型。 */
export class DocxComplexParseError extends Error {
  readonly code: DocxComplexParseErrorCode;
  readonly details: Record<string, unknown> | undefined;

  constructor(
    code: DocxComplexParseErrorCode,
    message: string,
    options: DocxComplexParseErrorOptions = {},
  ) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = 'DocxComplexParseError';
    this.code = code;
    this.details = options.details;
  }

  /** 可 JSON 序列化的形态（Error 自身的属性并不保证可序列化）。 */
  toJSON(): {
    code: DocxComplexParseErrorCode;
    message: string;
    details?: Record<string, unknown>;
  } {
    return this.details === undefined
      ? { code: this.code, message: this.message }
      : { code: this.code, message: this.message, details: this.details };
  }
}

/** 判断一个值是否为本模块错误。 */
export function isDocxComplexParseError(value: unknown): value is DocxComplexParseError {
  return value instanceof DocxComplexParseError;
}

/** 判断一个字符串是否是合法错误码。 */
export function isKnownErrorCode(code: string): code is DocxComplexParseErrorCode {
  return Object.prototype.hasOwnProperty.call(DOCX_COMPLEX_PARSE_ERROR_CODES, code);
}

/**
 * 把任意抛出物收敛为本模块错误。
 *
 * 用途只有一个：handler 的 `catch`。已经是我们自己的错误就原样放行（它的码是
 * 有意义的信息，重新包装只会把它埋掉）；其余一切（类型错误、`undefined` 访问、
 * 第三方库抛出的东西）才用 `fallbackCode` 兜底。
 */
export function toDocxComplexParseError(
  value: unknown,
  fallbackCode: DocxComplexParseErrorCode,
  context: { message?: string; path?: string } = {},
): DocxComplexParseError {
  if (isDocxComplexParseError(value)) return value;

  const reason = value instanceof Error ? value.message : String(value);
  const details: Record<string, unknown> = { reason };
  // 保留原始错误的 name：`TypeError` 与 `RangeError` 指向完全不同的排查方向，
  // 只留一句 message 会把这些线索丢掉。
  if (value instanceof Error && value.name) details['origin'] = value.name;
  if (context.path !== undefined) details['path'] = context.path;

  return new DocxComplexParseError(
    fallbackCode,
    context.message ?? 'Unexpected failure while handling a complex DOCX artifact',
    { details, cause: value },
  );
}

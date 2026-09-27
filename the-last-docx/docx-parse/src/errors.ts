/**
 * docx-parse 的错误类型。
 *
 * 离开本模块的每一个失败，都是 `DocxParseError`。它携带：
 *   - 稳定的、可编程判断的 `code`（机器可读，避免上层做字符串匹配）；
 *   - 可选的 `path`（类 JSON 指针，指向出问题的部件/字段）；
 *   - 可 JSON 序列化的 `details` 上下文。
 *
 * 之所以要自建错误类型而不直接抛引擎异常：
 *   1. 原生异常可能携带 Buffer、句柄等不可序列化内容，跨越 RPC 边界会失败；
 *   2. 上层需要按错误码分支，而不是解析错误消息文本；
 *   3. 换引擎时错误语义必须保持稳定。
 */
import type { DocxParseErrorCode, JsonObject, JsonValue } from './contract';

/** 构造 `DocxParseError` 的可选参数。 */
export interface DocxParseErrorOptions {
  /** 出问题的位置，例如 `word/document.xml`。 */
  path?: string;
  /** 附加上下文，必须可 JSON 序列化。 */
  details?: JsonObject;
  /** 原始异常，仅用于保留调用栈，不参与序列化。 */
  cause?: unknown;
}

/** 本模块唯一的错误类型。 */
export class DocxParseError extends Error {
  /** 稳定错误码。 */
  readonly code: DocxParseErrorCode;
  /** 出错位置（可选）。 */
  readonly path?: string;
  /** 结构化上下文（始终存在，最小为空对象）。 */
  readonly details: JsonObject;

  constructor(code: DocxParseErrorCode, message: string, options: DocxParseErrorOptions = {}) {
    // 注意：仅当 caller 显式传入 cause 时才把它挂到 Error 上，
    // 避免在未需要时创建无意义的包装链。
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = 'DocxParseError';
    this.code = code;
    if (options.path !== undefined) {
      this.path = options.path;
    }
    this.details = options.details ?? {};
  }

  /**
   * 稳定、JSON 安全的表示形式，用于日志、遥测与 RPC 边界。
   *
   * 之所以不用 `JSON.stringify(this)`：Error 的自有属性（message/stack）默认不可枚举，
   * 直接序列化会得到 `{}`。这里显式列字段，保证跨边界传输不丢信息。
   */
  toJSON(): JsonObject {
    const json: JsonObject = {
      name: this.name,
      code: this.code,
      message: this.message,
      details: this.details,
    };
    if (this.path !== undefined) {
      json.path = this.path;
    }
    return json;
  }
}

/** 类型守卫：判断一个未知值是否为 `DocxParseError`。 */
export function isDocxParseError(value: unknown): value is DocxParseError {
  return value instanceof DocxParseError;
}

/**
 * 把任意抛出物包装成 `DocxParseError`，且不丢失原始信息。
 *
 * 使用位置：每个公开 handler 的最外层。这样内部实现即使抛出原生异常
 * （例如 Node 的 `Error`、Python 桥接层的异常），对外也统一成模块错误。
 *
 * @param value        被抛出的原始值
 * @param fallbackCode 无法识别时使用的错误码
 * @param context      可选的补充消息与路径
 */
export function toDocxParseError(
  value: unknown,
  fallbackCode: DocxParseErrorCode,
  context: { message?: string; path?: string } = {},
): DocxParseError {
  // 已经是本模块错误则原样返回，避免重复包装导致 code 被降级覆盖。
  if (isDocxParseError(value)) {
    return value;
  }

  const detail: JsonObject = {};
  if (value instanceof Error) {
    detail.name = value.name;
    detail.reason = value.message;
  } else {
    detail.reason = stringifyUnknown(value);
  }

  return new DocxParseError(fallbackCode, context.message ?? 'docx-parse failed', {
    details: detail,
    path: context.path,
    cause: value,
  });
}

/**
 * 把任意值尽力转换为可 JSON 序列化的形态。
 *
 * 用途：错误上下文里可能出现函数、Symbol、BigInt 等不可序列化值，
 * 若不处理会导致错误对象本身在序列化时二次失败。
 */
export function stringifyUnknown(value: unknown): JsonValue {
  if (value === null) return null;
  switch (typeof value) {
    case 'string':
    case 'boolean':
      return value;
    case 'number':
      // NaN / Infinity 不是合法 JSON，转成字符串形式保留信息。
      return Number.isFinite(value) ? value : String(value);
    case 'bigint':
      // BigInt 无法直接序列化，降级为字符串。
      return value.toString();
    case 'undefined':
      return null;
    case 'object':
      if (Array.isArray(value)) {
        return value.map((entry) => stringifyUnknown(entry));
      }
      if (value instanceof Error) {
        return { name: value.name, message: value.message };
      }
      // 普通对象不做深拷贝：避免循环引用导致的无限递归。
      // 用类型标签保留足够诊断信息即可。
      return { reason: Object.prototype.toString.call(value) };
    default:
      // function / symbol
      return String(value);
  }
}

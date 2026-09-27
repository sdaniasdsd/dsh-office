/**
 * docx-parse 的遥测实现。
 *
 * 设计目标：
 *   1. 结构化：事件是普通对象，天然可 JSON 序列化，可直接投递给任意采集器；
 *   2. 隐私优先：事件里【绝不】写入文档正文、绝对路径等敏感信息，只记录计数与耗时；
 *   3. 可注入：Profile 可以提供自己的 `TelemetrySink`，本模块只依赖这个窄接口。
 */
import type { JsonValue, Logger, TelemetryEvent, TelemetryLevel, TelemetrySink } from './contract';

/**
 * 内存中保留的事件条数上限。
 *
 * 用环形缓冲（超出即丢最旧）而不是无界数组：独立运行时无人消费事件，
 * 若不设上限会随请求量无限增长，最终拖垮进程。
 */
const MAX_RETAINED_EVENTS = 256;

/** 跨度（span）的属性集合。 */
export interface TelemetrySpanAttributes {
  [key: string]: JsonValue;
}

/**
 * 默认的内存遥测实现。
 *
 * 典型用途：单元测试中断言「是否记录了引擎耗时」，
 * 或在没有采集器的独立运行场景下保留最近的事件便于排查。
 */
export class InMemoryTelemetry implements TelemetrySink {
  private readonly events: TelemetryEvent[] = [];
  private readonly now: () => number;

  /**
   * @param now 时间源，默认 `Date.now`。注入时间源可让测试得到确定性的耗时。
   */
  constructor(now: () => number = () => Date.now()) {
    this.now = now;
  }

  record(event: TelemetryEvent): void {
    this.events.push(event);
    if (this.events.length > MAX_RETAINED_EVENTS) {
      // 一次性裁掉多余部分，避免每个事件都做 splice(0,1)。
      this.events.splice(0, this.events.length - MAX_RETAINED_EVENTS);
    }
  }

  /** 返回事件的浅拷贝快照，调用方修改不会影响内部缓冲。 */
  snapshot(): TelemetryEvent[] {
    return this.events.map((event) => ({ ...event, attributes: { ...event.attributes } }));
  }

  /** 清空缓冲（测试常用）。 */
  clear(): void {
    this.events.length = 0;
  }

  /** 暴露内部时间源，保证外部测量与事件时间戳同源。 */
  nowMs(): number {
    return this.now();
  }
}

/** 丢弃一切事件的实现；当 Profile 注入自己的 sink 时作为默认值。 */
export const noopTelemetry: TelemetrySink = {
  record(): void {
    // 有意为空：上游已提供采集方案，本模块无需自行保留。
  },
};

/**
 * 执行 `fn` 并记录一对 start/end 跨度事件；失败时记录 error 事件后原样抛出。
 *
 * 注意：这里【重新抛出】原始错误，不做包装——错误分类是 `errors.ts` 的职责，
 * 遥测只负责观察，不改变控制流语义。
 *
 * @param sink       遥测接收端
 * @param name       事件名，建议形如 `docx-parse.<phase>`
 * @param attributes 起始即已知的属性（不要把结果类信息塞进来）
 * @param fn         被观测的异步操作
 */
export async function withSpan<T>(
  sink: TelemetrySink,
  name: string,
  attributes: TelemetrySpanAttributes,
  fn: () => Promise<T>,
): Promise<T> {
  const startedAt = Date.now();
  sink.record({
    name,
    level: 'debug',
    timestampMs: startedAt,
    attributes: { ...attributes, phase: 'start' },
  });
  try {
    const value = await fn();
    sink.record({
      name,
      level: 'debug',
      timestampMs: Date.now(),
      durationMs: Date.now() - startedAt,
      attributes: { ...attributes, phase: 'end' },
    });
    return value;
  } catch (error) {
    sink.record({
      name,
      level: 'error',
      timestampMs: Date.now(),
      durationMs: Date.now() - startedAt,
      attributes: {
        ...attributes,
        phase: 'error',
        // 只记录消息文本，不记录堆栈与路径，避免泄露环境信息。
        reason: error instanceof Error ? error.message : String(error),
      },
    });
    throw error;
  }
}

/** 低层事件写入 helper：统一补全时间戳与空属性。 */
function emit(
  sink: TelemetrySink | undefined,
  level: TelemetryLevel,
  name: string,
  attributes?: Record<string, JsonValue>,
): void {
  if (!sink) return;
  sink.record({ name, level, timestampMs: Date.now(), attributes: attributes ?? {} });
}

/**
 * 基于遥测的日志适配器。
 *
 * 之所以不直接 `console.log`：日志与遥测走同一条链路，便于统一采集与开关控制；
 * 同时避免污染 stdout（Python 桥接依赖 stdout 传协议数据，TS 侧也可能被宿主复用）。
 */
export function createTelemetryLogger(sink?: TelemetrySink): Logger {
  return {
    debug: (message, attributes) => emit(sink, 'debug', message, attributes),
    info: (message, attributes) => emit(sink, 'info', message, attributes),
    warn: (message, attributes) => emit(sink, 'warn', message, attributes),
    error: (message, attributes) => emit(sink, 'error', message, attributes),
  };
}

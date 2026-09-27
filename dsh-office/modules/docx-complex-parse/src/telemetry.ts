/**
 * 本模块的遥测实现。
 *
 * 设计目标：
 *   1. 结构化：事件是普通对象，天然可 JSON 序列化，可直接投递给任意采集器；
 *   2. 隐私优先：事件里【绝不】写入文档正文、绝对路径等敏感信息，只记录计数与耗时；
 *   3. 可注入：Profile 可以提供自己的 `TelemetrySink`，本模块只依赖这个窄接口。
 *
 * 注意本模块的 `TelemetryEvent` 里**没有时间戳字段**：时间是采集端的职责。
 * 模块内测量出的耗时（引擎毫秒、总毫秒）走 `TelemetrySummary` 回给调用方，
 * 而不是塞进事件属性——否则同一份事件在不同采集端会有两套时间语义。
 */
import type { JsonValue, Logger, TelemetryEvent, TelemetryLevel, TelemetrySink } from './contract';

/**
 * 内存中保留的事件条数上限。
 *
 * 用环形缓冲（超出即丢最旧）而不是无界数组：独立运行时若无人消费事件，
 * 不设上限会随请求量无限增长，最终拖垮进程。
 */
const MAX_RETAINED_EVENTS = 256;

/**
 * 默认的内存遥测实现。
 *
 * 典型用途：测试中断言「是否记录了引擎耗时」，或在没有采集器的独立运行场景下
 * 保留最近的事件便于排查。
 */
export class InMemoryTelemetry implements TelemetrySink {
  private readonly events: TelemetryEvent[] = [];

  emit(event: TelemetryEvent): void {
    this.events.push(event);
    if (this.events.length > MAX_RETAINED_EVENTS) {
      // 一次性裁掉多余部分，避免每个事件都做 splice(0, 1)。
      this.events.splice(0, this.events.length - MAX_RETAINED_EVENTS);
    }
  }

  /** 返回事件的浅拷贝快照，调用方修改不会影响内部缓冲。 */
  snapshot(): TelemetryEvent[] {
    return this.events.map((event) => ({ ...event, attributes: { ...event.attributes } }));
  }

  /** 按事件名筛选事件（测试常用）。 */
  eventsNamed(name: string): TelemetryEvent[] {
    return this.snapshot().filter((event) => event.name === name);
  }

  /** 已保留的事件条数。 */
  get size(): number {
    return this.events.length;
  }

  /** 清空缓冲（测试常用）。 */
  clear(): void {
    this.events.length = 0;
  }
}

/** 丢弃一切事件的实现；当 Profile 已提供自己的 sink 时作为默认值。 */
export const noopTelemetry: TelemetrySink = {
  emit(): void {
    // 有意为空：上游已有采集方案，本模块无需自行保留。
  },
};

/** 低层事件写入 helper：统一处理「未注入 sink」与空属性。 */
function emitEvent(
  sink: TelemetrySink | undefined,
  level: TelemetryLevel,
  name: string,
  attributes?: Record<string, JsonValue>,
): void {
  if (!sink) return;
  sink.emit({ level, name, attributes: attributes ?? {} });
}

/**
 * 基于遥测的日志适配器。
 *
 * 之所以不直接 `console.log`：日志与遥测走同一条链路，便于统一采集与开关控制；
 * 同时避免污染 stdout——Python 桥接依赖 stdout 传协议数据，宿主也可能复用 stdout。
 */
export function createTelemetryLogger(sink?: TelemetrySink): Logger {
  return {
    debug: (name, attributes) => emitEvent(sink, 'debug', name, attributes),
    info: (name, attributes) => emitEvent(sink, 'info', name, attributes),
    warn: (name, attributes) => emitEvent(sink, 'warn', name, attributes),
    error: (name, attributes) => emitEvent(sink, 'error', name, attributes),
  };
}

/**
 * 执行 `fn` 并记录一对 start/end 事件；失败时记录 error 事件后原样抛出。
 *
 * 注意：这里【重新抛出】原始错误，不做包装——错误分类是 `errors.ts` 的职责，
 * 遥测只负责观察，不改变控制流语义。
 */
export async function withSpan<T>(
  sink: TelemetrySink | undefined,
  name: string,
  attributes: Record<string, JsonValue>,
  fn: () => Promise<T>,
): Promise<T> {
  emitEvent(sink, 'debug', name, { ...attributes, phase: 'start' });
  try {
    const value = await fn();
    emitEvent(sink, 'debug', name, { ...attributes, phase: 'end' });
    return value;
  } catch (error) {
    emitEvent(sink, 'error', name, {
      ...attributes,
      phase: 'error',
      // 只记录消息文本，不记录堆栈与路径，避免泄露环境信息。
      reason: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }
}

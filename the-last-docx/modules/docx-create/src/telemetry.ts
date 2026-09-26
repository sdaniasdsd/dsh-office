import type { TelemetrySink } from './contract';
/** Diagnostics must not undo a successful artifact commit or leak document content. */
export function recordTelemetry(sink: TelemetrySink | undefined, event: Parameters<TelemetrySink['record']>[0]): void {
  try { void Promise.resolve(sink?.record(event)).catch(() => undefined); } catch { /* isolated */ }
}

import type { TelemetrySink } from './contract';

export function recordRenderTelemetry(
  sink: TelemetrySink | undefined,
  input: { operation: 'inspect' | 'execute' | 'verify'; requestId: string; engine: string; startedAt: number; pageCount?: number },
): void {
  sink?.record({
    name: `docx-render.${input.operation}`,
    timestampMs: input.startedAt,
    durationMs: Date.now() - input.startedAt,
    attributes: {
      engine: input.engine,
      requestId: input.requestId,
      ...(input.pageCount === undefined ? {} : { pageCount: input.pageCount }),
    },
  });
}

import type { ModuleCreateOptions } from './contract';
export function record(sink:ModuleCreateOptions['telemetry'],event:Parameters<NonNullable<ModuleCreateOptions['telemetry']>['record']>[0]):void {
  try { void Promise.resolve(sink?.record(event)).catch(()=>undefined); } catch { /* telemetry cannot invalidate a delivery */ }
}

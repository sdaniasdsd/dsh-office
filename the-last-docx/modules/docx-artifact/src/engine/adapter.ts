import type { ArtifactEngine, ModuleCreateOptions } from '../contract';
import { fail } from '../errors';
/** No storage implementation or renderer is duplicated here; both are host-owned ports. */
export function createEngine(options:ModuleCreateOptions):ArtifactEngine {
  if(options.engine) return options.engine;
  const files=options.files, inspector=options.inspector;
  if(!files||!inspector) fail('INVALID_INPUT','office-files and DOCX inspector adapters must be injected.');
  return {read:(ref,max,signal)=>files.read(ref,max,signal),inspect:(ref,signal)=>inspector.inspect(ref,signal),commit:input=>files.commitManifest(input),async dispose(){}};
}

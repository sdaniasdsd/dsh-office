import { resolve } from 'node:path';
import { LocalArtifactFiles } from './files';
import { DocxProfile } from './profile';
/** Environment is consumed ONLY by the plugin host, never by capability modules. */
export async function createHost(){
  const workspace=resolve(process.env.THE_LAST_DOCX_WORKSPACE??process.cwd());
  const data=resolve(process.env.THE_LAST_DOCX_DATA??resolve(workspace,'.docx-data'));
  const files=await LocalArtifactFiles.create(data,[workspace]);
  // Do not inject bare `python` / `soffice` / `pdftoppm` defaults here. A
  // defined engine path is an explicit override in each partition and would
  // hide its manifest-driven runtime-package resolver. Direct DOCX_* paths
  // remain supported by the resolver itself.
  const packageNames=process.env.DSH_OFFICE_RUNTIME_PACKAGES?.split(',').map(value=>value.trim()).filter(Boolean);
  return new DocxProfile({files,runtimeRoot:process.env.DSH_OFFICE_RUNTIME_ROOT,runtimePackageNames:packageNames});
}

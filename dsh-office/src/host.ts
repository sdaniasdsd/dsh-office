import { resolve } from 'node:path';
import { LocalArtifactFiles } from './files';
import { DocxProfile } from './profile';
/** Environment is consumed ONLY by the plugin host, never by capability modules. */
export async function createHost(){
  const workspace=resolve(process.env.THE_LAST_DOCX_WORKSPACE??process.cwd());
  const data=resolve(process.env.THE_LAST_DOCX_DATA??resolve(workspace,'.docx-data'));
  const files=await LocalArtifactFiles.create(data,[workspace]);
  return new DocxProfile({files,pythonPath:process.env.DOCX_PYTHON??'python',sofficePath:process.env.DOCX_SOFFICE??'soffice',pdftoppmPath:process.env.DOCX_PDFTOPPM??'pdftoppm'});
}

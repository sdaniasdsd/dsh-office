import { createHash } from 'node:crypto';
import { createDocxArtifactModule,DocxArtifactError } from '../src/index';
import type { ArtifactRef,ArtifactFiles,DeliveryManifest } from '../src/contract';
const hash=(b:Uint8Array)=>createHash('sha256').update(b).digest('hex');
export function fixture(){
  const data=new Map<string,Uint8Array>(),versions=new Map<string,ArtifactRef>();
  const put=(name:string,bytes:Uint8Array):ArtifactRef=>{data.set(name,bytes);return {id:name,uri:`memory://${name}`,sha256:hash(bytes),sizeBytes:bytes.length};};
  const doc=put('doc',new TextEncoder().encode('DOCX fixture; inspector is injected'));
  const files:ArtifactFiles={async read(ref){const bytes=data.get(ref.id);if(!bytes)throw new Error('missing');return bytes;},async commitManifest(input){
    const key=`${input.documentId}:${input.revision}`,existing=versions.get(key);
    if(existing){if(existing.sha256!==hash(input.bytes))throw new DocxArtifactError('VERSION_CONFLICT','Conflict');return existing;}
    if(input.revision>1&&versions.get(`${input.documentId}:${input.revision-1}`)?.sha256!==input.parentSha256)throw new DocxArtifactError('VERSION_CONFLICT','Missing parent');
    const ref=put(key,input.bytes);versions.set(key,ref);return ref;
  }};
  const inspector={async inspect(){return {format:'docx' as const,mediaType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',extension:'docx',confidence:1,container:'zip' as const,encrypted:false,signatures:['fixture'],features:{},metadata:{}};}};
  const module=createDocxArtifactModule({files,inspector});
  const request={artifactRef:doc,operation:'execute' as const,requestId:'test',delivery:{documentId:'report',revision:1}};
  const manifest=(ref:ArtifactRef)=>JSON.parse(new TextDecoder().decode(data.get(ref.id)!)) as DeliveryManifest;
  return {put,data,doc,files,inspector,module,request,manifest,versions};
}

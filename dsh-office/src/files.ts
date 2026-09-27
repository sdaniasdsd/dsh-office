import { createHash,randomUUID } from 'node:crypto';
import { open,mkdir,realpath,link,unlink } from 'node:fs/promises';
import { resolve,join,relative,isAbsolute,basename,extname } from 'node:path';
import { fileURLToPath,pathToFileURL } from 'node:url';
import type { ArtifactRef } from 'office-core';
import { DocxArtifactError } from '@dsh-office-profile/docx-artifact';
export const sha256=(bytes:Uint8Array)=>createHash('sha256').update(bytes).digest('hex');
const isWithin=(parent:string,path:string)=>{const r=relative(parent,path);return r===''||(!r.startsWith('..')&&!isAbsolute(r));};
function cancelled(signal?:AbortSignal){if(signal?.aborted)throw new DocxArtifactError('ENGINE_TIMEOUT','Operation cancelled before commit.');}
const mime=(name:string)=>({'.docx':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','.pptx':'application/vnd.openxmlformats-officedocument.presentationml.presentation','.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','.pdf':'application/pdf','.png':'image/png','.json':'application/json'}[extname(name).toLowerCase()]??'application/octet-stream');
/** One host adapter accepts all EXISTING source / sources writer shapes without modifying them. */
export class LocalArtifactFiles {
  private constructor(readonly root:string,readonly inputRoots:string[]){}
  static async create(root:string,inputRoots:string[]){
    await mkdir(root,{recursive:true});const canonical=await realpath(root);
    const roots=await Promise.all(inputRoots.map(p=>realpath(p)));return new LocalArtifactFiles(canonical,roots);
  }
  private path(ref:ArtifactRef){if(!ref.uri.startsWith('file:'))throw new DocxArtifactError('ARTIFACT_UNAVAILABLE','Only managed local file references are accepted.');return fileURLToPath(ref.uri);}
  private async bounded(path:string,max:number){
    const f=await open(path,'r');try{
      const info=await f.stat();if(!info.isFile()||info.size>max)throw new DocxArtifactError('LIMIT_EXCEEDED','File exceeds its byte budget.');
      const bytes=new Uint8Array(info.size+1);let offset=0;
      while(offset<bytes.length){const {bytesRead}=await f.read(bytes,offset,bytes.length-offset,offset);if(!bytesRead)break;offset+=bytesRead;}
      if(offset!==info.size)throw new DocxArtifactError('INTEGRITY_MISMATCH','File changed during read.');return bytes.subarray(0,offset);
    }finally{await f.close();}
  }
  async read(ref:ArtifactRef,maxBytes=64*1024*1024,signal?:AbortSignal):Promise<Uint8Array>{
    cancelled(signal);const path=await realpath(this.path(ref));
    if(!isWithin(this.root,path))throw new DocxArtifactError('SAFETY_POLICY_DENIED','Reference is outside the managed artifact store. Import it first.');
    const bytes=await this.bounded(path,maxBytes),hash=sha256(bytes);
    if(ref.id!==`sha256:${hash}`||(ref.sha256&&ref.sha256!==hash)||(ref.sizeBytes!==undefined&&ref.sizeBytes!==bytes.length))throw new DocxArtifactError('INTEGRITY_MISMATCH','Managed artifact identity or metadata is stale.');
    cancelled(signal);return bytes;
  }
  async importFile(path:string,maxBytes=64*1024*1024):Promise<ArtifactRef>{
    const canonical=await realpath(path.startsWith('file:')?fileURLToPath(path):resolve(path));
    if(!this.inputRoots.some(root=>isWithin(root,canonical)))throw new DocxArtifactError('SAFETY_POLICY_DENIED','Import is outside configured input roots.');
    return this.write({bytes:await this.bounded(canonical,maxBytes),suggestedName:basename(canonical)});
  }
  private async publish(path:string,bytes:Uint8Array,signal?:AbortSignal){
    const dir=resolve(path,'..');await mkdir(dir,{recursive:true});
    const actual=await realpath(dir);if(!isWithin(this.root,actual))throw new DocxArtifactError('SAFETY_POLICY_DENIED','Store directory escapes its root.');
    const temp=join(actual,`.pending-${randomUUID()}`),file=await open(temp,'wx');
    try{
      try{await file.writeFile(bytes);await file.sync();}finally{await file.close();}
      try{cancelled(signal);await link(temp,path);}catch(error){
      if((error as NodeJS.ErrnoException).code!=='EEXIST')throw error;
      const existingPath=await realpath(path);if(!isWithin(this.root,existingPath))throw new DocxArtifactError('SAFETY_POLICY_DENIED','Existing store entry escapes its root.');
      let existing:Uint8Array;
      try{existing=await this.bounded(existingPath,bytes.length);}catch{throw new DocxArtifactError('VERSION_CONFLICT','Immutable artifact has different bytes or size.');}
      if(sha256(existing)!==sha256(bytes))throw new DocxArtifactError('VERSION_CONFLICT','Immutable artifact already contains different bytes.');
      }
    }finally{await unlink(temp).catch(()=>undefined);}
  }
  private ref(path:string,bytes:Uint8Array,label:string):ArtifactRef{const hash=sha256(bytes);return {id:`sha256:${hash}`,uri:pathToFileURL(path).href,sha256:hash,sizeBytes:bytes.length,label,mediaType:mime(label)};}
  async write(input:{bytes:Uint8Array;suggestedName:string;requestId?:string;source?:ArtifactRef;sources?:ArtifactRef[]}):Promise<ArtifactRef>{
    const label=basename(input.suggestedName).replace(/[^\p{L}\p{N}._-]/gu,'_').slice(0,120)||'artifact.bin';
    const path=join(this.root,'objects',sha256(input.bytes),`artifact${extname(label).toLowerCase()}`);await this.publish(path,input.bytes);return this.ref(path,input.bytes,label);
  }
  async materialize(ref:ArtifactRef){await this.read(ref);return {path:this.path(ref)};}
  async commitManifest(input:{documentId:string;revision:number;parentSha256?:string;requestId:string;bytes:Uint8Array;signal:AbortSignal}):Promise<ArtifactRef>{
    if(!/^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/.test(input.documentId)||!Number.isSafeInteger(input.revision)||input.revision<1)throw new DocxArtifactError('INVALID_INPUT','Invalid document version key.');
    const dir=join(this.root,'versions',`doc-${input.documentId}`);await mkdir(dir,{recursive:true});
    const canonical=await realpath(dir);if(!isWithin(this.root,canonical))throw new DocxArtifactError('SAFETY_POLICY_DENIED','Version directory escapes its root.');
    // A version is write-once. Atomic exclusive linking is the CAS; there is no mutable
    // HEAD and no process lock file that could remain stale after a crash.
      cancelled(input.signal);
      if(input.revision>1){const prior=join(canonical,`${input.revision-1}.json`);let bytes;
        try{const priorPath=await realpath(prior);if(!isWithin(this.root,priorPath))throw new Error('Escaped version path');bytes=await this.bounded(priorPath,64*1024*1024);}catch{throw new DocxArtifactError('VERSION_CONFLICT','Previous revision is not committed in this store.');}
        if(sha256(bytes)!==input.parentSha256)throw new DocxArtifactError('VERSION_CONFLICT','Parent digest does not match committed history.');
      }else if(input.parentSha256)throw new DocxArtifactError('VERSION_CONFLICT','First revision cannot have a parent.');
      const path=join(canonical,`${input.revision}.json`);await this.publish(path,input.bytes,input.signal);return this.ref(path,input.bytes,`${input.documentId}.v${input.revision}.manifest.json`);
  }
  async readJson(ref:ArtifactRef){return JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(await this.read(ref)));}
}

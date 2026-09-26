import { createHash } from 'node:crypto';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { ArtifactRef, DeliveryManifest, DocxArtifactModule, ModuleCreateOptions, ModuleConfig, ModuleOutput, VerificationReport } from './contract';
import { createEngine } from './engine/adapter';
import { configSchema, resolveConfig } from './config';
import { parse, inspectSchema, executeSchema, verifySchema, manifestSchema, refSchema, canonicalJson } from './domain/docx-artifact';
import { fail, normalized, DocxArtifactError } from './errors';
import { toManifest } from './mapper';
import { validateReferences, reportFor } from './verifier';
import { record } from './telemetry';
export const DOCX_ARTIFACT_SCHEMAS=Object.fromEntries(Object.entries({config:configSchema,inspect:inspectSchema,execute:executeSchema,verify:verifySchema,manifest:manifestSchema}).map(([k,v])=>[k,zodToJsonSchema(v,{$refStrategy:'none'})]));
export const DOCX_ARTIFACT_DEFINITION={id:'docx-artifact',version:'0.1.0',profileGroup:'DOCX',capabilities:['inspect','execute','verify'] as const,
  configSchema:DOCX_ARTIFACT_SCHEMAS.config!,dependencies:['office-core','office-safety','office-files','office-preview','office-test-kit'].map(name=>({name,kind:'module'}))};
const digest=(bytes:Uint8Array)=>createHash('sha256').update(bytes).digest('hex');
export const MANIFEST_MIME='application/vnd.dsh.docx-delivery+json';
export function createDocxArtifactModule(options:ModuleCreateOptions):DocxArtifactModule {
  const base=resolveConfig(options.config),engine=createEngine(options);let disposed=false;
  async function run<T>(operation:string,task:(signal:AbortSignal)=>Promise<T>,timeout:number):Promise<T> {
    const start=Date.now(),controller=new AbortController();let code:string|undefined;
    const timer=setTimeout(()=>controller.abort(),timeout);
    try { if(disposed) fail('MODULE_DISPOSED','Delivery module is disposed.');
      return await Promise.race([task(controller.signal),new Promise<never>((_,reject)=>controller.signal.addEventListener('abort',()=>reject(new DocxArtifactError('ENGINE_TIMEOUT','Delivery exceeded its time budget.')),{once:true}))]);
    } catch(error) {const e=normalized(error);code=e.code;throw e;}
    finally {clearTimeout(timer);record(options.telemetry,{moduleId:'docx-artifact',operation,ok:!code,elapsedMs:Date.now()-start,...(code?{errorCode:code}:{})});}
  }
  function session(config:ModuleConfig,signal:AbortSignal,operation:string,policy?:Parameters<NonNullable<ModuleCreateOptions['safetyGuard']>['assertAllowed']>[0]['policy']) {
    let count=0,total=0;
    const identities=new Map<string,string>();
    async function read(ref:ArtifactRef,max=config.limits.maxArtifactBytes):Promise<{ref:ArtifactRef;bytes:Uint8Array}> {
      if(signal.aborted||disposed) fail('ENGINE_TIMEOUT','Delivery was cancelled before I/O.');
      if(++count>config.limits.maxReferences) fail('LIMIT_EXCEEDED','Too many delivery references.');
      try {await options.safetyGuard?.assertAllowed({artifactRef:ref,operation,...(policy?{policy}:{})});}catch {fail('SAFETY_POLICY_DENIED','Delivery safety guard denied a reference.');}
      let bytes:Uint8Array;
      try {bytes=await engine.read(ref,max,signal);}catch(e){if(e instanceof DocxArtifactError)throw e;return fail('ARTIFACT_UNAVAILABLE','A delivery reference could not be read.');}
      if(!(bytes instanceof Uint8Array)||bytes.length>max||(total+=bytes.length)>config.limits.maxTotalBytes) fail('LIMIT_EXCEEDED','Delivery bytes exceed configured budget.');
      const hash=digest(bytes);
      if((ref.sha256&&ref.sha256!==hash)||(ref.sizeBytes!==undefined&&ref.sizeBytes!==bytes.length)) fail('INTEGRITY_MISMATCH','Artifact metadata disagrees with its bytes.');
      if(identities.has(ref.id)&&identities.get(ref.id)!==hash) fail('INTEGRITY_MISMATCH','One artifact ID refers to different content.');
      identities.set(ref.id,hash);
      return {ref:{...ref,sha256:hash,sizeBytes:bytes.length},bytes};
    }
    async function readManifest(ref:ArtifactRef) {
      const loaded=await read(ref,config.limits.maxManifestBytes);
      let data:unknown;try {data=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(loaded.bytes));}catch{return fail('FORMAT_MISMATCH','Malformed delivery manifest.');}
      return {ref:loaded.ref,manifest:parse(manifestSchema,data)};
    }
    async function validate(manifest:DeliveryManifest):Promise<DeliveryManifest> {
      const m=parse(manifestSchema,manifest);m.document=(await read(m.document)).ref;
      const profile=await engine.inspect(m.document,signal);
      if(profile.format!=='docx'||profile.encrypted!==false) fail('FORMAT_MISMATCH','Delivery requires an inspected unencrypted DOCX.');
      validateReferences(m,config);
      if(m.preview){
        if(m.preview.pdf){const item=await read(m.preview.pdf);if(Buffer.from(item.bytes.subarray(0,5)).toString()!=='%PDF-')fail('FORMAT_MISMATCH','Preview PDF has an invalid signature.');m.preview.pdf={...item.ref,mediaType:'application/pdf'};}
        for(const page of m.preview.pages) for(const key of ['image','thumbnail'] as const) if(page[key]) {
          const item=await read(page[key]!);
          if(!Buffer.from(item.bytes.subarray(0,8)).equals(Buffer.from([137,80,78,71,13,10,26,10]))) fail('FORMAT_MISMATCH','Page preview must be a PNG.');
          page[key]={...item.ref,mediaType:'image/png'};
        }
      }
      for(const bridge of m.bridges){
        const item=await read(bridge.artifactRef);let evidence:unknown;
        try{evidence=JSON.parse(new TextDecoder('utf-8',{fatal:true}).decode(item.bytes));}catch{return fail('FORMAT_MISMATCH','Bridge evidence must be a JSON envelope.');}
        const e=evidence as {schema?:unknown;kind?:unknown;source?:{id?:unknown;sha256?:unknown}};
        if(e?.schema!=='docx-evidence/v1'||e.kind!==bridge.kind||e.source?.id!==m.document.id||e.source.sha256!==m.document.sha256)fail('STALE_REFERENCE','Stored bridge evidence is not bound to this document revision.');
        bridge.artifactRef={...item.ref,mediaType:'application/json'};
      }
      let current=m;
      while(current.parentManifest){
        const parent=await readManifest(current.parentManifest);
        if(parent.manifest.documentId!==m.documentId||parent.manifest.revision!==current.revision-1)fail('VERSION_CONFLICT','Parent belongs to a different document or nonadjacent revision.');
        if(current===m)m.parentManifest=parent.ref;
        current=parent.manifest;
      }
      if(current.revision!==1)fail('VERSION_CONFLICT','Version chain does not terminate at revision 1.');
      return m;
    }
    return {read,readManifest,validate};
  }
  const output=<T>(operation:ModuleOutput<T>['operation'],requestId:string,result:T):ModuleOutput<T>=>({moduleId:'docx-artifact',operation,requestId,result,artifacts:[],warnings:[]});
  return {definition:DOCX_ARTIFACT_DEFINITION,handlers:{
    async inspect(raw){const input=parse(inspectSchema,raw),config=resolveConfig(input.options,base,true);return run('inspect',async signal=>{
      const s=session(config,signal,'inspect',input.policy),doc=await s.read(input.artifactRef),profile=await engine.inspect(doc.ref,signal);
      if(profile.format!=='docx'||profile.encrypted!==false)fail('FORMAT_MISMATCH','Expected an unencrypted DOCX.');
      return output('inspect',input.requestId,profile);
    },config.timeoutMs);},
    async execute(raw){const input=parse(executeSchema,raw),config=resolveConfig(input.options,base,true);return run('execute',async signal=>{
      const s=session(config,signal,'execute',input.policy),m=await s.validate(toManifest(input.artifactRef,input.delivery));
      const bytes=new TextEncoder().encode(canonicalJson(m));if(bytes.length>config.limits.maxManifestBytes)fail('LIMIT_EXCEEDED','Manifest exceeds its byte budget.');
      if(signal.aborted||disposed)fail('ENGINE_TIMEOUT','Cancelled before manifest commit.');
      const rawRef=await engine.commit({documentId:m.documentId,revision:m.revision,...(m.parentManifest?{parentSha256:m.parentManifest.sha256!}:{}),requestId:input.requestId,bytes,signal});
      const checked=parse(refSchema,rawRef);
      if((checked.sha256&&checked.sha256!==digest(bytes))||(checked.sizeBytes!==undefined&&checked.sizeBytes!==bytes.length)||checked.id===m.document.id||checked.uri===m.document.uri)fail('INTEGRITY_MISMATCH','Store returned an invalid manifest reference.');
      const manifestRef={...checked,sha256:digest(bytes),sizeBytes:bytes.length,mediaType:MANIFEST_MIME};
      return {...output('execute',input.requestId,{manifestRef,documentRef:m.document,manifest:m}),artifacts:[manifestRef],verification:reportFor(m),warnings:m.review.state==='pending'?[{code:'VISUAL_REVIEW_PENDING',severity:'info' as const,message:'Delivery references verified; visual review has not been supplied.'}]:[]};
    },config.timeoutMs);},
    async verify(raw){const input=parse(verifySchema,raw),config=resolveConfig(input.options,base,true);return run('verify',async signal=>{
      let report:VerificationReport;
      try{const s=session(config,signal,'verify',input.policy),loaded=await s.readManifest(input.artifactRef);report=reportFor(await s.validate(loaded.manifest));}
      catch(error){const e=normalized(error);if(['SAFETY_POLICY_DENIED','ENGINE_TIMEOUT'].includes(e.code))throw e;report={ok:false,partial:true,checks:[{id:'delivery',status:'fail',severity:'error',message:`${e.code}: ${e.message}`}],summary:{total:1,passed:0,failed:1,skipped:0}};}
      return {...output('verify',input.requestId,report),verification:report};
    },config.timeoutMs);}
  },async dispose(){if(!disposed){disposed=true;await engine.dispose();}}};
}
export * from './contract';
export { DocxArtifactError, ERROR_CODES } from './errors';
export { manifestSchema,canonicalJson } from './domain/docx-artifact';
export async function register(registry:{registerModule(module:DocxArtifactModule):void|Promise<void>},options:ModuleCreateOptions){const module=createDocxArtifactModule(options);await registry.registerModule(module);return module;}

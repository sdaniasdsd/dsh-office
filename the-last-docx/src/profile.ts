import { createDocxInspectModule } from '@dsh-office-profile/docx-inspect';
import { createDocxParseModule as createDual } from '@dsh-office-profile/docx-parse';
import { createDocxParseModule as createEasy,resolveConfig as easyConfig } from '@dsh-office-profile/docx-easy-parse';
import { createDocxComplexParseModule } from '@dsh-office-profile/docx-complex-parse';
import { createDocxCreateModule } from '@dsh-office-profile/docx-create';
import { createDocxStylesModule } from '@dsh-office-profile/docx-styles';
import { createDocxEditModule } from '@dsh-office-profile/docx-edit';
import { createDocxRenderModule } from '@dsh-office-profile/docx-render';
import { createDocxArtifactModule,canonicalJson,DocxArtifactError } from '@dsh-office-profile/docx-artifact';
import type { BridgeKind,BridgeReference } from '@dsh-office-profile/docx-artifact';
import type { ArtifactRef } from 'office-core';
import { LocalArtifactFiles } from './files';
import { bridgeComplexToDual } from './source-bridge';
export const MODULE_IDS=['docx-inspect','docx-easy-parse','docx-parse','docx-complex-parse','docx-create','docx-styles','docx-edit','docx-render','docx-artifact'] as const;
export type ModuleId=typeof MODULE_IDS[number];
/** The modules whose execute path is gated by a pre-flight policy. */
export type PolicyGatedModuleId='docx-edit'|'docx-render'|'docx-artifact';
/**
 * The pre-flight policy the Profile applies before a module writes, renders or
 * delivers.
 *
 * `docx-edit` allows a document that already carries an external link. An
 * in-place edit writes text and properties; it never creates, drops or
 * retargets a relationship, and `docx-edit` now proves exactly that after the
 * write by comparing the external-relationship set. Refusing the document up
 * front also refused the ordinary case - a project report with a runbook link
 * that only needed one number changed.
 *
 * Rendering and delivery stay strict: a renderer or a recipient may resolve
 * links on its own.
 */
export const DEFAULT_EXECUTE_POLICIES:Record<PolicyGatedModuleId,Record<string,unknown>>={
  'docx-edit':{id:'profile-default-edit',allowMacros:false,allowExternalLinks:true,allowEmbeddedObjects:false,allowAltChunks:false,allowEncrypted:false},
  'docx-render':{id:'profile-default-render',allowMacros:false,allowExternalLinks:false,allowEmbeddedObjects:false,allowAltChunks:false,allowEncrypted:false},
  'docx-artifact':{id:'profile-default-delivery',allowMacros:false,allowExternalLinks:false,allowEmbeddedObjects:false,allowAltChunks:false,allowEncrypted:false},
};
export interface ProfileOptions {
  files:LocalArtifactFiles; pythonPath?:string; sofficePath?:string; pdftoppmPath?:string;
  /** Trusted Profile-level configuration only; never accepted from a document or tool request. */
  modules?:Partial<Record<ModuleId,Record<string,unknown>>>;
  /** Trusted override of {@link DEFAULT_EXECUTE_POLICIES}, per gated module. */
  safetyPolicy?:Partial<Record<PolicyGatedModuleId,Record<string,unknown>>>;
}
type JsonRecord=Record<string,unknown>;
type ModulePort={definition:{id:string;[key:string]:unknown};handlers:Record<string,(input:JsonRecord)=>Promise<JsonRecord>>;dispose?:()=>Promise<void>};
export interface RegisteredDocxModule {
  definition:ModulePort['definition'];
  handlers:Record<'inspect'|'execute'|'verify',(input:JsonRecord)=>Promise<JsonRecord>>;
}
export interface ProfileRegistryLike {registerModule(module:RegisteredDocxModule):void|Promise<void>}
export class DocxProfile {
  private instances=new Map<ModuleId,ModulePort>();private closed=false;
  constructor(readonly options:ProfileOptions){}
  get files(){return this.options.files;}
  private get(id:ModuleId):ModulePort{
    if(this.closed)throw new DocxArtifactError('MODULE_DISPOSED','Profile is closed.');
    const found=this.instances.get(id);if(found)return found;
    const python={pythonPath:this.options.pythonPath??'python'};
    const config=this.options.modules?.[id]??{};
    let instance:unknown;
    switch(id){
      case 'docx-inspect':instance=createDocxInspectModule({config:{...config,engine:{...python,...(config.engine as object??{})}}});break;
      case 'docx-parse':instance=createDual({config:{...config,engine:{...python,...(config.engine as object??{})}}});break;
      case 'docx-easy-parse':instance=createEasy({config:easyConfig({...config,engine:{...python,...(config.engine as object??{})}})});break;
      case 'docx-complex-parse':instance=createDocxComplexParseModule({config:{...config,engine:{...python,...(config.engine as object??{})}}});break;
      case 'docx-create':instance=createDocxCreateModule({artifactStore:this.files,config});break;
      case 'docx-styles':instance=createDocxStylesModule({artifactStore:this.files,config});break;
      case 'docx-edit':instance=createDocxEditModule({artifactStore:this.files,config});break;
      case 'docx-render':instance=createDocxRenderModule({artifactStore:this.files,config:{...config,engine:{sofficePath:this.options.sofficePath??'soffice',pdftoppmPath:this.options.pdftoppmPath??'pdftoppm',...(config.engine as object??{})}}});break;
      case 'docx-artifact':instance=createDocxArtifactModule({files:this.files,config,inspector:{inspect:async ref=>{
        const result=await this.call('docx-inspect','inspect',{artifactRef:ref,requestId:'delivery-inspection'});return result.result as import('office-core').FormatProfile;
      }}});break;
    }
    const port=instance as ModulePort;this.instances.set(id,port);return port;
  }
  list(){return MODULE_IDS.map(id=>({registryId:id,definition:this.get(id).definition,
    ...(id==='docx-easy-parse'?{compatibility:{originalModuleId:'docx-parse',resultShape:'legacy-easy-ir',note:'Source ID and output retained; only registry/package name is distinct.'}}:{})}));}
  /** The host owns lifecycle. Register once; the target registry owns duplicate handling. */
  async registerWith(registry:ProfileRegistryLike):Promise<void>{
    for(const id of MODULE_IDS){
      const native=this.get(id);
      await registry.registerModule({definition:{...native.definition,id},handlers:{
        inspect:input=>this.call(id,'inspect',input),
        execute:input=>this.call(id,'execute',input),
        verify:input=>this.call(id,'verify',input),
      }});
    }
  }
  async call(id:ModuleId,operation:'inspect'|'execute'|'verify',request:JsonRecord):Promise<JsonRecord>{
    if(!MODULE_IDS.includes(id)||!['inspect','execute','verify'].includes(operation))throw new DocxArtifactError('INVALID_INPUT','Unknown module or capability.');
    const cloned=JSON.parse(canonicalJson(request)) as JsonRecord;
    if(typeof cloned.requestId!=='string'||!cloned.requestId.trim())throw new DocxArtifactError('INVALID_INPUT','requestId is required.');
    const opts=cloned.options as JsonRecord|undefined;
    if(opts?.engine||opts?.limits||opts?.timeoutMs)throw new DocxArtifactError('SAFETY_POLICY_DENIED','Runtime paths and resource budgets are configured by Profile, not by tool calls.');
    if((opts?.featureFlags as JsonRecord|undefined)?.enforceLimits===false)throw new DocxArtifactError('SAFETY_POLICY_DENIED','Plugin calls cannot disable safety limits.');
    // Every reference is materialized/verified by office-files before a path-based legacy engine sees it.
    const visit=async(value:unknown):Promise<void>=>{
      if(!value||typeof value!=='object')return;
      if(!Array.isArray(value)&&typeof (value as JsonRecord).uri==='string'&&typeof (value as JsonRecord).id==='string')await this.files.read(value as ArtifactRef);
      for(const child of Object.values(value))await visit(child);
    };
    await visit(cloned);cloned.operation=operation;
    const module=this.get(id);
    // Safety policy stays owned by the inspect module; validate source before write/render/delivery.
    if(['docx-edit','docx-render','docx-artifact'].includes(id)&&operation==='execute'&&cloned.artifactRef){
      const gated=id as PolicyGatedModuleId;
      const policy={...DEFAULT_EXECUTE_POLICIES[gated],...(this.options.safetyPolicy?.[gated]??{})};
      await this.call('docx-inspect','inspect',{artifactRef:cloned.artifactRef,requestId:`${cloned.requestId}:safety`,policy});
    }
    const result=await module.handlers[operation]!(cloned);canonicalJson(result);return result;
  }
  async evidence(kind:BridgeKind,source:ArtifactRef,payload:unknown):Promise<BridgeReference>{
    await this.files.read(source);if(!source.sha256)throw new DocxArtifactError('INTEGRITY_MISMATCH','Import source before recording evidence.');
    const bytes=new TextEncoder().encode(canonicalJson({schema:'docx-evidence/v1',kind,source:{id:source.id,sha256:source.sha256},payload}));
    const ref=await this.files.write({bytes,suggestedName:`${kind}.json`});
    return {kind,artifactRef:ref,sourceArtifactId:source.id,sourceSha256:source.sha256};
  }
  async analyze(source:ArtifactRef,requestId:string,complex:'off'|'structure'|'layout'='off',easy=false){
    const references:BridgeReference[]=[];
    const inspected=await this.call('docx-inspect','execute',{artifactRef:source,requestId:`${requestId}:inspect`});references.push(await this.evidence('inspection',source,inspected));
    const dual=await this.call('docx-parse','execute',{artifactRef:source,requestId:`${requestId}:dual`});references.push(await this.evidence('dual-ir',source,dual));
    if(easy)references.push(await this.evidence('easy-ir',source,await this.call('docx-easy-parse','execute',{artifactRef:source,requestId:`${requestId}:easy`})));
    if(complex!=='off'){
      const result=await this.call('docx-complex-parse','execute',{artifactRef:source,requestId:`${requestId}:complex`,options:{featureFlags:{parsePageGeometry:complex==='layout'}}});
      references.push(await this.evidence('complex-ir',source,result));
      references.push(await this.evidence('node-bridge',source,bridgeComplexToDual(dual.result as Parameters<typeof bridgeComplexToDual>[0],result.result as Parameters<typeof bridgeComplexToDual>[1])));
    }
    return {artifactRef:source,bridges:references,complexMode:complex};
  }
  async dispose(){if(this.closed)return;this.closed=true;const instances=[...this.instances.values()].reverse();this.instances.clear();await Promise.allSettled(instances.map(m=>m.dispose?.()));}
}

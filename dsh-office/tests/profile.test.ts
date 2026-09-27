import { describe,it,expect } from 'vitest';
import type { ModuleOutput as CreateOutput,CreateResult } from '@dsh-office-profile/docx-create';
import type { ExecuteResult as ParseResult } from '@dsh-office-profile/docx-parse';
import { targetFromDualIR } from '@dsh-office-profile/docx-edit';
import type { DeliveryResult } from '@dsh-office-profile/docx-artifact';
import type { ArtifactRef } from 'office-core';
import { context,createRequest } from './support';
import { docxParse,type RegisteredDocxModule } from '../src/index';
import { MODULE_IDS } from '../src/profile';
describe('all-module integration',()=>{
  it('registers unique host adapters and keeps the dual-IR middleware exports',async()=>{
    const c=await context();try{
      const registered:RegisteredDocxModule[]=[];
      await c.profile.registerWith({registerModule:module=>{registered.push(module);}});
      expect(registered.map(m=>m.definition.id)).toEqual(c.profile.list().map(m=>m.registryId));
      const created=await registered.find(m=>m.definition.id==='docx-create')!.handlers.execute(createRequest);
      expect(created.moduleId).toBe('docx-create');
      for(const api of ['createRevisionLedger','diffDualIR','irFingerprint','decideGovernance','renderAgentBrief'] as const)expect(docxParse[api]).toBeTypeOf('function');
      await c.profile.dispose();
      await expect(registered[0]!.handlers.inspect({requestId:'closed'})).rejects.toMatchObject({code:'MODULE_DISPOSED'});
    }finally{await c.close();}
  });
  it('registers every module the profile declares, and no duplicates',async()=>{
    const c=await context();try{
      const list=c.profile.list();
      // Asserted against the declaration rather than a literal count, so adding
      // a module cannot leave a stale number behind the way `8` did.
      expect(new Set(list.map(m=>m.registryId)).size).toBe(MODULE_IDS.length);
      expect(list.map(m=>m.registryId)).toEqual([...MODULE_IDS]);
      expect(list.find(m=>m.registryId==='docx-easy-parse')?.definition.id).toBe('docx-parse');
    }finally{await c.close();}
  });
  it('create -> inspect/easy/dual/complex -> bridge -> edit -> deliver -> verify -> next version',async()=>{
    const c=await context();try{
      const created=await c.profile.call('docx-create','execute',createRequest) as unknown as CreateOutput<CreateResult>;
      const source=created.result.artifactRef;
      const analyzed=await c.profile.analyze(source,'analysis','structure',true);
      expect(analyzed.bridges.map(b=>b.kind)).toEqual(['inspection','dual-ir','easy-ir','complex-ir','node-bridge']);
      const bridge=await c.files.readJson(analyzed.bridges.find(b=>b.kind==='node-bridge')!.artifactRef);
      expect(bridge.payload.coverage.matched).toBeGreaterThan(0);
      const raw=await c.files.readJson(analyzed.bridges.find(b=>b.kind==='dual-ir')!.artifactRef);
      const parsed=raw.payload.result as ParseResult;
      const paragraph=parsed.ir.content.semantic.blocks.find(b=>b.kind==='paragraph'&&b.anchor.quote==='初始内容');
      if(!paragraph)throw new Error('Body paragraph missing');
      const edited=await c.profile.call('docx-edit','execute',{requestId:'edit',artifactRef:source,plan:{edits:[{kind:'replaceText',target:targetFromDualIR(parsed.ir.content,paragraph.id),find:'初始内容',replace:'修改后的内容'}]}});
      const nextSource=(edited.artifacts as ArtifactRef[])[0]!;
      const reparsed=await c.profile.call('docx-parse','execute',{requestId:'reparse',artifactRef:nextSource});
      expect((reparsed.result as ParseResult).ir.content.semantic.blocks.find(b=>b.id===paragraph.id)?.anchor.quote).toBe('修改后的内容');
      const first=await c.profile.call('docx-artifact','execute',{requestId:'delivery',artifactRef:source,delivery:{documentId:'project',revision:1,bridges:analyzed.bridges}});
      const delivered=first.result as DeliveryResult;
      expect((await c.profile.call('docx-artifact','verify',{requestId:'verify',artifactRef:delivered.manifestRef})).result).toMatchObject({ok:true,partial:true});
      await expect(c.profile.call('docx-artifact','execute',{requestId:'stale',artifactRef:nextSource,delivery:{documentId:'project',revision:2,parentManifest:delivered.manifestRef,bridges:analyzed.bridges}})).rejects.toMatchObject({code:'STALE_REFERENCE'});
      const next=await c.profile.call('docx-artifact','execute',{requestId:'v2',artifactRef:nextSource,delivery:{documentId:'project',revision:2,parentManifest:delivered.manifestRef}});
      expect((next.result as DeliveryResult).manifest.revision).toBe(2);
      expect((await c.profile.call('docx-artifact','verify',{requestId:'verify-v2',artifactRef:(next.result as DeliveryResult).manifestRef})).result).toMatchObject({ok:true,partial:true});
    }finally{await c.close();}
  });
});

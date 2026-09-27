import { describe,it,expect } from 'vitest';
import { fixture } from './support';
import type { RenderResult } from '../src/contract';
describe('preview and evidence binding',()=>{
  function preview(f:ReturnType<typeof fixture>):RenderResult{return {engine:'fixture',sourceArtifactId:f.doc.id,sourceSha256:f.doc.sha256,pageCount:1,visualReview:'pending',findings:[],pages:[{pageNumber:1,image:f.put('image',new Uint8Array([137,80,78,71,13,10,26,10]))}]};}
  it('rejects stale preview and invalid page numbering',async()=>{
    const f=fixture(),p=preview(f);p.sourceSha256='0'.repeat(64);
    await expect(f.module.handlers.execute({...f.request,delivery:{...f.request.delivery,preview:p}})).rejects.toMatchObject({code:'STALE_REFERENCE'});
    p.sourceSha256=f.doc.sha256;p.pages[0]!.pageNumber=2;
    await expect(f.module.handlers.execute({...f.request,delivery:{...f.request.delivery,preview:p}})).rejects.toMatchObject({code:'INVALID_INPUT'});
  });
  it('does not promote pending review but accepts explicit reviewed findings',async()=>{
    const f=fixture(),p=preview(f);
    const out=await f.module.handlers.execute({...f.request,delivery:{...f.request.delivery,preview:p}});expect(out.verification?.partial).toBe(true);
    p.visualReview='provided';
    const reviewed=await f.module.handlers.execute({...f.request,delivery:{documentId:'reviewed',revision:1,preview:p}});expect(reviewed.result.manifest.review).toMatchObject({state:'reviewed',basis:'caller-provided-findings'});
    p.findings=[{id:'f',kind:'overlap',severity:'error',message:'overlap',page:1}];
    await expect(f.module.handlers.execute({...f.request,delivery:{documentId:'failed',revision:1,preview:p}})).rejects.toMatchObject({code:'VERIFICATION_FAILED'});
  });
  it('reads source binding from stored evidence rather than trusting the outer reference',async()=>{
    const f=fixture(),ref=f.put('evidence',new TextEncoder().encode(JSON.stringify({schema:'docx-evidence/v1',kind:'dual-ir',source:{id:'another',sha256:f.doc.sha256},payload:{}})));
    await expect(f.module.handlers.execute({...f.request,delivery:{...f.request.delivery,bridges:[{kind:'dual-ir',artifactRef:ref,sourceArtifactId:f.doc.id,sourceSha256:f.doc.sha256!}]}})).rejects.toMatchObject({code:'STALE_REFERENCE'});
  });
});

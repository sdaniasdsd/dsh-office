import { describe,it,expect } from 'vitest';
import { fixture } from './support';
import { DOCX_ARTIFACT_SCHEMAS,DocxArtifactError } from '../src/index';
describe('delivery contract',()=>{
  it('inspects and delivers references without claiming visual success',async()=>{
    const f=fixture(),out=await f.module.handlers.execute(f.request);
    expect(out.result.manifest.document.sha256).toBe(f.doc.sha256);
    expect(out.verification).toMatchObject({ok:true,partial:true});
    expect(JSON.parse(JSON.stringify(out))).toEqual(out);
    expect((await f.module.handlers.inspect({artifactRef:f.doc,operation:'inspect',requestId:'inspect'})).result.format).toBe('docx');
    expect((await f.module.handlers.verify({artifactRef:out.result.manifestRef,operation:'verify',requestId:'verify'})).result).toMatchObject({ok:true,partial:true});
  });
  it('commits immutable idempotent versions with an explicit prior manifest',async()=>{
    const f=fixture(),first=await f.module.handlers.execute(f.request),retry=await f.module.handlers.execute({...f.request,requestId:'retry'});
    expect(retry.result.manifestRef).toEqual(first.result.manifestRef);
    const next=await f.module.handlers.execute({...f.request,delivery:{...f.request.delivery,revision:2,parentManifest:first.result.manifestRef}});
    expect(next.result.manifest.parentManifest?.sha256).toBe(first.result.manifestRef.sha256);
    expect((await f.module.handlers.verify({artifactRef:next.result.manifestRef,requestId:'v2',operation:'verify'})).result.ok).toBe(true);
  });
  it('serializes stable errors and exposes machine-readable schemas',()=>{
    expect(JSON.parse(JSON.stringify(new DocxArtifactError('STALE_REFERENCE','stale'))).code).toBe('STALE_REFERENCE');
    expect(DOCX_ARTIFACT_SCHEMAS.manifest).toHaveProperty('properties');
  });
});

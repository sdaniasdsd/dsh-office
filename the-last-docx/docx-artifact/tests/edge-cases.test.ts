import { describe,it,expect,vi } from 'vitest';
import { fixture } from './support';
import { createDocxArtifactModule } from '../src/index';
describe('delivery rejection',()=>{
  it('rejects hash tampering and missing files before commit',async()=>{
    const f=fixture(),commit=vi.spyOn(f.files,'commitManifest');f.data.set('doc',new Uint8Array([1]));
    await expect(f.module.handlers.execute(f.request)).rejects.toMatchObject({code:'INTEGRITY_MISMATCH'});expect(commit).not.toHaveBeenCalled();
  });
  it('rejects invalid versions, missing parents, unsafe keys and configuration escalation',async()=>{
    const f=fixture();
    await expect(f.module.handlers.execute({...f.request,delivery:{documentId:'../escape',revision:1}})).rejects.toMatchObject({code:'INVALID_INPUT'});
    await expect(f.module.handlers.execute({...f.request,delivery:{documentId:'report',revision:2}})).rejects.toMatchObject({code:'VERSION_CONFLICT'});
    await expect(f.module.handlers.execute({...f.request,options:{featureFlags:{allowFailedReview:true}}})).rejects.toMatchObject({code:'SAFETY_POLICY_DENIED'});
  });
  it('applies guard, bytes limit and preview requirement',async()=>{
    const f=fixture();
    const denied=createDocxArtifactModule({files:f.files,inspector:f.inspector,safetyGuard:{async assertAllowed(){throw new Error('denied');}}});
    await expect(denied.handlers.execute(f.request)).rejects.toMatchObject({code:'SAFETY_POLICY_DENIED'});
    await expect(f.module.handlers.execute({...f.request,options:{limits:{maxArtifactBytes:1}}})).rejects.toMatchObject({code:'LIMIT_EXCEEDED'});
    await expect(f.module.handlers.execute({...f.request,options:{featureFlags:{requirePreview:true}}})).rejects.toMatchObject({code:'VERIFICATION_FAILED'});
  });
  it('verify returns a failed report for missing evidence, not an empty successful manifest',async()=>{
    const f=fixture(),out=await f.module.handlers.execute(f.request);f.data.delete('doc');
    const verified=await f.module.handlers.verify({operation:'verify',requestId:'verify',artifactRef:out.result.manifestRef});expect(verified.result).toMatchObject({ok:false,partial:true});
  });
  it('rejects disposed calls',async()=>{const f=fixture();await f.module.dispose();await expect(f.module.handlers.execute(f.request)).rejects.toMatchObject({code:'MODULE_DISPOSED'});});
});

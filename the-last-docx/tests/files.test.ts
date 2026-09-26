import { describe,it,expect } from 'vitest';
import { writeFile,mkdir,symlink,readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { context,createRequest } from './support';
import type { ArtifactRef } from 'office-core';
import type { DeliveryResult } from '@dsh-office-profile/docx-artifact';
describe('file and version boundaries',()=>{
  it('atomically selects one concurrent commit and cleans cancelled staging files',async()=>{
    const c=await context();try{
      const commit=(text:string,signal=new AbortController().signal)=>c.files.commitManifest({documentId:'CON',revision:1,requestId:'race',bytes:new TextEncoder().encode(text),signal});
      const race=await Promise.allSettled([commit('first'),commit('second')]);
      expect(race.filter(r=>r.status==='fulfilled')).toHaveLength(1);
      expect(race.find(r=>r.status==='rejected')).toMatchObject({reason:{code:'VERSION_CONFLICT'}});
      expect(await readdir(join(c.files.root,'versions','doc-CON'))).toEqual(['1.json']);
      const controller=new AbortController();controller.abort();
      await expect(commit('third',controller.signal)).rejects.toMatchObject({code:'ENGINE_TIMEOUT'});
      expect(await readdir(join(c.files.root,'versions','doc-CON'))).toEqual(['1.json']);
    }finally{await c.close();}
  });
  it('imports a snapshot and refuses traversal, external URIs and stale bytes',async()=>{
    const c=await context();try{
      const path=join(c.inputs,'source.txt');await writeFile(path,'original');const ref=await c.files.importFile(path);await writeFile(path,'changed');expect(new TextDecoder().decode(await c.files.read(ref))).toBe('original');
      const outside=join(c.dir,'outside.txt');await writeFile(outside,'outside');await expect(c.files.importFile(outside)).rejects.toMatchObject({code:'SAFETY_POLICY_DENIED'});
      await expect(c.files.read({...ref,uri:'https://example.com/document'})).rejects.toMatchObject({code:'ARTIFACT_UNAVAILABLE'});
      await writeFile(fileURLToPath(ref.uri),'tampered');await expect(c.files.read(ref)).rejects.toMatchObject({code:'INTEGRITY_MISMATCH'});
    }finally{await c.close();}
  });
  it('checks real paths through directory junctions',async()=>{
    const c=await context();try{
      const outside=join(c.dir,'outside');await mkdir(outside);await writeFile(join(outside,'secret.txt'),'not allowed');await symlink(outside,join(c.inputs,'junction'),'junction');
      await expect(c.files.importFile(join(c.inputs,'junction','secret.txt'))).rejects.toMatchObject({code:'SAFETY_POLICY_DENIED'});
    }finally{await c.close();}
  });
  it('same revision is idempotent; competing content cannot overwrite it',async()=>{
    const c=await context();try{
      const created=await c.profile.call('docx-create','execute',createRequest),doc=(created.result as {artifactRef:ArtifactRef}).artifactRef;
      const request={requestId:'delivery',artifactRef:doc,delivery:{documentId:'versions',revision:1}};
      const first=await c.profile.call('docx-artifact','execute',request),second=await c.profile.call('docx-artifact','execute',request);
      expect((first.result as DeliveryResult).manifestRef).toEqual((second.result as DeliveryResult).manifestRef);
      const changed=await c.profile.call('docx-create','execute',{requestId:'changed',plan:{kind:'create',document:{blocks:[{kind:'paragraph',id:'new',runs:[{text:'other'}]}]}}});
      await expect(c.profile.call('docx-artifact','execute',{...request,artifactRef:(changed.result as {artifactRef:ArtifactRef}).artifactRef})).rejects.toMatchObject({code:'VERSION_CONFLICT'});
      expect((await c.profile.call('docx-artifact','verify',{requestId:'still-valid',artifactRef:(first.result as DeliveryResult).manifestRef})).result).toMatchObject({ok:true});
    }finally{await c.close();}
  });
  it('rejects untrusted attempts to disable limits or change executable paths',async()=>{
    const c=await context();try{await expect(c.profile.call('docx-inspect','execute',{requestId:'unsafe',options:{engine:{pythonPath:'other'}}})).rejects.toMatchObject({code:'SAFETY_POLICY_DENIED'});}finally{await c.close();}
  });
});

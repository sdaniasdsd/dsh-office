import { mkdtemp,mkdir,rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { LocalArtifactFiles } from '../src/files';
import { DocxProfile } from '../src/profile';
export async function context(){const dir=await mkdtemp(join(tmpdir(),'last-docx-test-')),inputs=join(dir,'input');await mkdir(inputs);const files=await LocalArtifactFiles.create(join(dir,'data'),[inputs]);const profile=new DocxProfile({files});return {dir,inputs,files,profile,async close(){await profile.dispose();await rm(dir,{recursive:true,force:true});}};}
export const createRequest={requestId:'create',plan:{kind:'create',document:{header:'项目报告',blocks:[{kind:'paragraph',id:'title',style:'Heading1',runs:[{text:'测试标题'}]},{kind:'paragraph',id:'body',runs:[{text:'初始内容'}]},{kind:'table',id:'table',rows:[['项目','说明'],['模块','测试']]}]}}};

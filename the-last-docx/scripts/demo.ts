import { mkdir,writeFile } from 'node:fs/promises';
import { resolve,join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { LocalArtifactFiles,DocxProfile } from '../src/index';
import type { ArtifactRef } from 'office-core';
import type { DeliveryResult } from '@dsh-office-profile/docx-artifact';
const root=fileURLToPath(new URL('..',import.meta.url)),out=resolve(root,'examples','generated',randomUUID());await mkdir(out,{recursive:true});
const files=await LocalArtifactFiles.create(join(out,'store'),[out]),profile=new DocxProfile({files});
try{
  const created=await profile.call('docx-create','execute',{requestId:'demo-create',plan:{kind:'create',document:{preset:'report',header:'The Last DOCX · 集成交付示例',footer:'结构验证完成，视觉排版待验收',blocks:[
    {kind:'paragraph',id:'title',style:'Title',runs:[{text:'统一文档交付示例'}]},
    {kind:'paragraph',id:'summary',runs:[{text:'本文件由 docx-create 创建，并通过检查、双 IR、轻量解析和复杂结构解析。'}]},
    {kind:'table',id:'modules',rows:[['模块','职责'],['解析','保持双 IR 和稳定编辑锚点'],['编辑','按节点修改，不覆盖原文'],['渲染','独立排版，不伪造视觉通过'],['交付','校验摘要、来源引用与版本链']]},
  ]}}});
  const document=(created.result as {artifactRef:ArtifactRef}).artifactRef,analysis=await profile.analyze(document,'demo-analysis','structure',true);
  const output=await profile.call('docx-artifact','execute',{requestId:'demo-delivery',artifactRef:document,delivery:{documentId:'integration-demo',revision:1,bridges:analysis.bridges}});
  const result=output.result as DeliveryResult,verification=await profile.call('docx-artifact','verify',{requestId:'demo-verify',artifactRef:result.manifestRef});
  await writeFile(join(out,'delivery-result.json'),JSON.stringify({result,verification},null,2),{flag:'wx'});
  console.log(JSON.stringify({directory:out,document:result.documentRef,manifest:result.manifestRef,verification:verification.result},null,2));
}finally{await profile.dispose();}

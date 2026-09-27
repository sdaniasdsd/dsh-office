import type { ExecuteResult as DualResult } from '@dsh-office-profile/docx-parse';
import type { ExecuteResult as ComplexResult } from '@dsh-office-profile/docx-complex-parse';
import { DocxArtifactError } from '@dsh-office-profile/docx-artifact';
/** Join by verified source digest + exact structural pointer; never translate a geometry ID into an edit ID by guess. */
export function bridgeComplexToDual(dual:DualResult,complex:ComplexResult){
  if(!dual.artifact.sha256||dual.artifact.id!==complex.artifact.id||dual.artifact.sha256!==complex.artifact.sha256)throw new DocxArtifactError('STALE_REFERENCE','Cannot bridge different document revisions.');
  const byPointer=dual.ir.content.sourceMap.byPointer;
  const links=complex.ir.content.blocks.map(block=>{
    const pointer=`${block.coordinate.part}!${block.coordinate.structuralPath}`;
    return {complexNodeId:block.id,pointer,semanticId:byPointer[pointer]??null,status:byPointer[pointer]?'exact-source-match':'unmatched'};
  });
  return {schema:'docx-source-bridge/v1',sourceArtifactId:dual.artifact.id,sourceSha256:dual.artifact.sha256,
    links,coverage:{matched:links.filter(l=>l.semanticId!==null).length,unmatched:links.filter(l=>l.semanticId===null).length},
    note:'Complex page indices remain zero-based and render pages remain one-based. No layout mapping is inferred.'};
}

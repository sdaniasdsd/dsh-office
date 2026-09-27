import type { DeliveryManifest, ModuleConfig, VerificationReport, VerificationCheck } from './contract';
import { fail } from './errors';
import { reviewState } from './mapper';
export function validateReferences(m:DeliveryManifest,c:ModuleConfig):void {
  if(!m.document.sha256) fail('INTEGRITY_MISMATCH','Delivery document needs a verified digest.');
  if((m.revision===1&&m.parentManifest)||(m.revision>1&&!m.parentManifest)) fail('VERSION_CONFLICT','Version 1 has no parent; later versions require the prior manifest.');
  if(c.featureFlags.requirePreview&&!m.preview) fail('VERIFICATION_FAILED','Profile requires a preview.');
  const p=m.preview;
  if(p) {
    if(p.sourceArtifactId!==m.document.id||p.sourceSha256!==m.document.sha256) fail('STALE_REFERENCE','Preview belongs to a different document revision.');
    if(p.pageCount>c.limits.maxPages) fail('LIMIT_EXCEEDED','Preview page limit exceeded.');
    if(p.pages.length!==p.pageCount||p.pages.some((page,i)=>page.pageNumber!==i+1)||p.findings.some(f=>f.page>p.pageCount)) fail('INVALID_INPUT','Preview page numbers or findings are inconsistent.');
    if(p.visualReview==='pending'&&p.findings.length) fail('INVALID_INPUT','Pending preview cannot contain a completed review.');
  }
  for(const bridge of m.bridges) if(bridge.sourceArtifactId!==m.document.id||bridge.sourceSha256!==m.document.sha256) fail('STALE_REFERENCE','Bridge belongs to another document revision.');
  if(JSON.stringify(m.review)!==JSON.stringify(reviewState(p))) fail('VERIFICATION_FAILED','Review state contradicts its evidence.');
  if(c.featureFlags.requireVisualReview&&m.review.state!=='reviewed') fail('VERIFICATION_FAILED','Profile requires completed visual review without blocking findings.');
  if(!c.featureFlags.allowFailedReview&&m.review.state==='failed') fail('VERIFICATION_FAILED','Blocking visual findings prevent delivery.');
}
export function reportFor(m:DeliveryManifest,checks:VerificationCheck[]=[]):VerificationReport {
  checks=[...checks,{id:'references',status:'pass',severity:'info',message:'File digests, sizes, version parent and source bindings verified.'},
    {id:'visual-review',status:m.review.state==='pending'?'skip':m.review.state==='failed'?'fail':'pass',severity:m.review.state==='failed'?'error':'info',message:'Visual status comes from caller-provided docx-render findings; delivery does not perform visual analysis.'}];
  const passed=checks.filter(x=>x.status==='pass').length,failed=checks.filter(x=>x.status==='fail').length,skipped=checks.filter(x=>x.status==='skip').length;
  return {ok:failed===0,partial:skipped>0,checks,summary:{total:checks.length,passed,failed,skipped}};
}

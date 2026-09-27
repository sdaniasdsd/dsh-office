import type { DeliveryManifest, DeliveryPlan, ArtifactRef } from './contract';
export function reviewState(preview:DeliveryPlan['preview']):DeliveryManifest['review'] {
  if(!preview||preview.visualReview!=='provided') return {state:'pending',basis:'none'};
  return {state:preview.findings.some(f=>f.severity==='error')?'failed':'reviewed',basis:'caller-provided-findings'};
}
export function toManifest(document:ArtifactRef,plan:DeliveryPlan):DeliveryManifest {
  return {schema:'docx-delivery/v1',documentId:plan.documentId,revision:plan.revision,document,
    ...(plan.parentManifest?{parentManifest:plan.parentManifest}:{}),...(plan.preview?{preview:plan.preview}:{}),bridges:plan.bridges??[],review:reviewState(plan.preview)};
}

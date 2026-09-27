# D09 — DSH Office Profile

## Verdict
First-pass: **PASS**, 93/100. Final: **PASS**, 93/100 after 0/5 repairs. One-step pass: **true**. Converged: **true**.

## Test contract
- Frozen primary prompt: “这份MHRA applicant response下周要给主管过目，现在标题和章节看起来像一整块。帮我把主标题、主要小节做出清楚但克制的层级，留白舒服一点，用稳重的蓝灰色；正文、条文、页码和表格里的字一个都别改。”
- Initial content: mhra-applicant-response.docx from the frozen official-source office corpus.
- Expected effects: 
  - Title and section hierarchy is visibly clearer.
  - Text, values, labels, legal/safety wording, and order remain unchanged.
- Invariants: 
  - Preserve source wording, facts, fields, values, and order.
  - Keep all pages and page numbering; no clipping or new blank pages.
  - Use DSH in-place formatting only; do not rebuild the source from a plan.
- Repair limit: 5.
- Acceptance gates: file integrity, complete render, content preservation, no critical visual defect, intent complete; threshold 85/100.

## Environment and reproducibility
- Runner: DSH Office Profile direct local module calls at the current repository revision; frozen natural-language prompt was manually mapped to the typed module operation (not autonomous prompt interpretation).
- Module: docx-edit; renderer: DSH docx-render / bundled LibreOffice; 120 DPI.
- Baseline renderer shared with candidate.
- Initial action result: module call and verifier completed.

## Artifact inventory and baseline visual inventory
- Source: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-stress-20260927-visual-confirm-200\workspace\docx\mhra-applicant-response.docx; SHA-256 53e04952c2c9b0b48063e70ad8f586b40e004a36957a0b898f2b4d0fefe43cff; unchanged: true.
- Baseline: 10 pages; 10 full-resolution images under pages; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\shared-baselines\docx\mhra-applicant-response\render\contact-01.png.
- Round 00 candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\artifact.docx; SHA-256 e5cd5dccd78b061fecfe1af20f409c3276c6d0b9c86ead9099b196eba98094ea; 10 pages; image inventory 10; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\render\contact-01.png.
- Final candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\artifact.docx; 10 page images; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\render\contact-01.png.
- Verifier evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\manifest.json; passed.

## Round log and scores
| Round | Gates | One-step | Intent | Craft | Thoroughness | Strength | Total |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | all pass | 30 | 23 | 17 | 15 | 8 | 93 |

### Round 00
- Exact prompt: “这份MHRA applicant response下周要给主管过目，现在标题和章节看起来像一整块。帮我把主标题、主要小节做出清楚但克制的层级，留白舒服一点，用稳重的蓝灰色；正文、条文、页码和表格里的字一个都别改。”
- Action: {"edits":[{"kind":"formatParagraph","target":{"semanticId":"530a4bb04ddb323b5b745309db8d9b00","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[2]","ordinal":1,"paraId":"5DC94A55","quote":"The applicant’s <Multidisciplinary>/<Quality>/<Clinical>/<Non-Clinical>/<Module 1> response","digest":"946406a33cec553915830f9726a9c6d881ac5e31b736c5e16c114bdefe2b4f04"}},"alignment":"Centered","spaceAfter":12,"font":{"bold":true,"size":21,"color":"#244A67","name":"Arial"}}]}
- Visual review: every rendered page inspected in the complete contact sheet; full-resolution page images retained. Evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\render\contact-01.png; all 10 page images retained.
- Gates: {"fileIntegrity":true,"completeRender":true,"contentPreservation":true,"noCriticalVisualDefect":true,"intentComplete":true}.

## Findings
- [minor] all 10 rendered pages reviewed via the contact sheet; no clipped or missing content found. Full-resolution pages are in D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D09\dsh\round-00\render\pages.

## Repair history and regressions
- None attempted; no initial artifact failure or repair was necessary.

## Final conclusion
**PASS** No failed gates. 

# D08 — DSH Office Profile

## Verdict
First-pass: **FAIL**, 20/100. Final: **FAIL**, 20/100 after 1/5 repair. One-step pass: **false**. Converged: **false**.

## Test contract
- Frozen primary prompt: “准备把flexible-working application form放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。”
- Initial content: flexible-working-form.docx from the frozen official-source office corpus.
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
- Initial action result: VERIFICATION_FAILED: The saved document did not pass post-edit verification.

## Artifact inventory and baseline visual inventory
- Source: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-stress-20260927-visual-confirm-200\workspace\docx\flexible-working-form.docx; SHA-256 f085f7328f2cb272c059744ee78a6b979e5a458eeef55859e3ecf3c754c72c62; unchanged: true.
- Baseline: 3 pages; 3 full-resolution images under pages; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\shared-baselines\docx\flexible-working-form\render\contact-01.png.
- Round 00 candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\artifact.docx; SHA-256 f085f7328f2cb272c059744ee78a6b979e5a458eeef55859e3ecf3c754c72c62; 3 pages; image inventory 3; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\render\contact-01.png.
- Final candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\artifact.docx; 3 page images; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\render\contact-01.png.
- Verifier evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\manifest.json; failed or unavailable.

## Round log and scores
| Round | Gates | One-step | Intent | Craft | Thoroughness | Strength | Total |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | fileIntegrity, contentPreservation, intentComplete | 0 | 0 | 10 | 10 | 0 | 20 |
| 1 | repair failed | 0 | 0 | 10 | 10 | 0 | 20 |

### Round 00
- Exact prompt: “准备把flexible-working application form放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。”
- Action: {"edits":[{"kind":"formatParagraph","target":{"semanticId":"4e3022ec1f2b878db99405c1a327d0a4","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[4]","ordinal":3,"paraId":"58435ED7","quote":"You can use this form to make an application to work flexibly under the right provided in law to eligible employees. Before completing this form, read the guida","digest":"b573ed00a30bf4806b061330483b2218419da452382d22e7d75cb35db4802182"}},"alignment":"Centered","spaceAfter":12,"font":{"bold":true,"size":21,"color":"#23415E","name":"Arial"}},{"kind":"formatParagraph","target":{"semanticId":"cf9fd107c299eb60d7e398f1383cceda","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[15]","ordinal":14,"paraId":"3281DD04","quote":"Name:","digest":"d370ae02b7c5b15d3905acf4461957cdc688950ee2fdb1d469ce638156d4ca27"}},"spaceBefore":6,"spaceAfter":3,"font":{"bold":true,"size":11,"color":"#23415E","name":"Arial"}},{"kind":"formatParagraph","target":{"semanticId":"430c61a9819279ca616f925568716719","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[16]","ordinal":15,"paraId":"3D1F568F","quote":"Staff or payroll number:","digest":"d94da9c92d95f0188c392dcfecdd463833f4604a97b0a94ca95f6aaf03603e44"}},"spaceBefore":6,"spaceAfter":3,"font":{"bold":true,"size":11,"color":"#23415E","name":"Arial"}},{"kind":"formatParagraph","target":{"semanticId":"c5e9c76a3e7f044723e4f7fcc8a8558b","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[17]","ordinal":16,"paraId":"564F5BE5","quote":"Manager:","digest":"256e95f4db9b3b4fedeb68353d2cfd49b6ad2ca57982ce2713a8a8e167d39682"}},"spaceBefore":6,"spaceAfter":3,"font":{"bold":true,"size":11,"color":"#23415E","name":"Arial"}},{"kind":"formatParagraph","target":{"semanticId":"bf6c3e188bc912b19eb1a1456a96cf7a","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[18]","ordinal":17,"paraId":"5587DA3B","quote":"National Insurance number:","digest":"f7f4b27dc9cfeb67692ac37a3fdbfea5b22ada4a22b4ebdb3a22aa6327cb84b2"}},"spaceBefore":6,"spaceAfter":3,"font":{"bold":true,"size":11,"color":"#23415E","name":"Arial"}},{"kind":"formatParagraph","target":{"semanticId":"0291027077f3342b3a2f53078fade91f","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[34]","ordinal":33,"paraId":"7D4D93DD","quote":"Date:","digest":"9f1a9aca974cd99b943ca8be2b5d6fd77eaa1dbb3e588ea36bbd51105ca867b7"}},"spaceBefore":6,"spaceAfter":3,"font":{"bold":true,"size":11,"color":"#23415E","name":"Arial"}}]}
- Visual review: every rendered page inspected in the complete contact sheet; full-resolution page images retained. Evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\render\contact-01.png; all 3 page images retained.
- Gates: {"fileIntegrity":false,"completeRender":true,"contentPreservation":false,"noCriticalVisualDefect":true,"intentComplete":false}.

### Round 01 repair
- Repair prompt: “准备把flexible-working application form放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。
Repair based on round-00 evidence: first call failed post-edit verification with generic diagnostics; reduce request to a directly-verifiable bold/color hierarchy only, preserving every word.”
- Diagnostic and changes: first call failed post-edit verification with generic diagnostics; reduce request to a directly-verifiable bold/color hierarchy only, preserving every word.
- DSH action: failed; output D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-01\artifact.docx; render 3 pages.
- Visual evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-01\render\contact-01.png; all 3 full-page images retained.
- Repair verdict: repair did not resolve the diagnosed blocker; round-00 remains the best candidate.

## Findings
- [major] round-00 render 3 pages; source identity retained after VERIFICATION_FAILED (The saved document did not pass post-edit verification.). Requested effect is absent; see complete contact sheet D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D08\dsh\round-00\render\contact-01.png.

## Repair history and regressions
- Round 1/5: first call failed post-edit verification with generic diagnostics; reduce request to a directly-verifiable bold/color hierarchy only, preserving every word. Result: still failed; unresolved.

## Final conclusion
**FAIL** — at least one hard gate failed or the requested effect was unavailable. Failed gates: fileIntegrity, contentPreservation, intentComplete, score_threshold. Repair was evidence-led and limited to one retry because the same verification class persisted.

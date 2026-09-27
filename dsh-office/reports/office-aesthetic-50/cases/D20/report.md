# D20 — DSH Office Profile

## Verdict
First-pass: **FAIL**, 20/100. Final: **FAIL**, 20/100 after 1/5 repair. One-step pass: **false**. Converged: **false**.

## Test contract
- Frozen primary prompt: “准备把QA evidence report放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。”
- Initial content: qa-evidence-report-template.docx from the frozen official-source office corpus.
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
- Initial action result: ENGINE_FAILED: Value [Center] is invalid for the enum type [org_docx4j_wml.JcEnumeration].

## Artifact inventory and baseline visual inventory
- Source: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-stress-20260927-visual-confirm-200\workspace\docx\qa-evidence-report-template.docx; SHA-256 f9774c99c1310122a059035ffd527333313b53734337dc03a4a114337c075090; unchanged: true.
- Baseline: 2 pages; 2 full-resolution images under pages; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\shared-baselines\docx\qa-evidence-report-template\render\contact-01.png.
- Round 00 candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\artifact.docx; SHA-256 f9774c99c1310122a059035ffd527333313b53734337dc03a4a114337c075090; 2 pages; image inventory 2; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\render\contact-01.png.
- Final candidate: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\artifact.docx; 2 page images; contact D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\render\contact-01.png.
- Verifier evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\manifest.json; failed or unavailable.

## Round log and scores
| Round | Gates | One-step | Intent | Craft | Thoroughness | Strength | Total |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | fileIntegrity, contentPreservation, intentComplete | 0 | 0 | 10 | 10 | 0 | 20 |
| 1 | repair failed | 0 | 0 | 10 | 10 | 0 | 20 |

### Round 00
- Exact prompt: “准备把QA evidence report放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。”
- Action: {"edits":[{"kind":"formatParagraph","target":{"semanticId":"2723cd539f5e45b4646e9f4fb2340959","anchor":{"kind":"p","part":"word/document.xml","structuralPath":"/w:document/w:body/w:p[1]","ordinal":0,"paraId":"157A3DD5","quote":"Analysis and Evidence Quality Assurance (QA) Report","digest":"029f02619a2223fe049871bbe73062eb93161a24d3677c1d3a4bea8952b18228"}},"alignment":"Centered","spaceAfter":12,"font":{"bold":true,"size":21,"color":"#23415E","name":"Arial"}},{"kind":"formatTable","target":{"semanticId":"b5d5fca2b6a5d801ac4220b83200e0ae","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[1]","ordinal":0,"paraId":null,"quote":"Name of product:","digest":"d0c79f122836588e66297fe4cd82b078e23c64b971a444d3b81f0193ec10aef8"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}},{"kind":"formatTable","target":{"semanticId":"7a58873c3af1b6c0794a869127971405","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[2]","ordinal":1,"paraId":null,"quote":"This template is to provide a record of peer review outcomes, together with clearance and approval decisions. This is to support the Department for Energy Secur","digest":"c72e222e21ab09212750d373304f0f6a69f2d5f8d7d652c21cd75375a81061e2"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}},{"kind":"formatTable","target":{"semanticId":"aedc240f78f289364450a07d5c9afbd3","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[3]","ordinal":2,"paraId":null,"quote":"RoleName & contact details Commissioner Analyst Assurer [Name, team]QA plans approved? (select)Date:[planned or actual] Analytical Clearance given? (select)Date","digest":"71fe0904668f9e33b613fed843e53c8cf12e511eeb63c3ccc20efff5dbf071f8"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}},{"kind":"formatTable","target":{"semanticId":"10a377aef628f7795a7c88108fa126a4","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[4]","ordinal":3,"paraId":null,"quote":"Please attach or link here evidence (confirmatory email/meeting minutes) of approval decision and caveats/recommendations, clarifications, and follow-up actions","digest":"406cd2d1219019bd4d828fd53710667d17fdf16e156a6af06639b620df437fe8"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}},{"kind":"formatTable","target":{"semanticId":"dd19da80865dfd02cec81f235d6c341f","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[5]","ordinal":4,"paraId":null,"quote":"Please provide a recommendation with supporting explanation to the Commissioner or Approver on the overall fitness for purpose of the analysis and evidence. Thi","digest":"bc69c034345615eb5850905b148c9c1e252fc428b9490b70cabe4738c22e31b6"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}},{"kind":"formatTable","target":{"semanticId":"e6a96c9ac07caaed4163c3477fe3d637","anchor":{"kind":"tbl","part":"word/document.xml","structuralPath":"/w:document/w:body/w:tbl[6]","ordinal":5,"paraId":null,"quote":"Name & contact details:(of Assurer or Reviewer)Date agreed: Review scope: What has and has not been reviewed. What QA did the Reviewer request? Review findings:","digest":"99609cabe3011da87532a32249cf832890d10a6f5d00c3ddb63910f7559c5e09"}},"layout":"autofit","alignment":"Center","cellVerticalAlignment":"center","cellMargins":{"top":3,"left":4,"bottom":3,"right":4},"borders":{"top":{"style":"single","size":0.5,"color":"#C8D2DC"},"bottom":{"style":"single","size":0.5,"color":"#C8D2DC"},"left":{"style":"single","size":0.5,"color":"#D6DEE5"},"right":{"style":"single","size":0.5,"color":"#D6DEE5"},"insideH":{"style":"single","size":0.35,"color":"#D6DEE5"},"insideV":{"style":"single","size":0.35,"color":"#D6DEE5"}}}]}
- Visual review: every rendered page inspected in the complete contact sheet; full-resolution page images retained. Evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\render\contact-01.png; all 2 page images retained.
- Gates: {"fileIntegrity":false,"completeRender":true,"contentPreservation":false,"noCriticalVisualDefect":true,"intentComplete":false}.

### Round 01 repair
- Repair prompt: “准备把QA evidence report放进正式评审包，麻烦做一轮统一的轻量美化：标题、章节、小表格的视觉语言协调，正文仍然严肃易读；所有法律/安全/审批措辞、字段、数字和顺序保持不变。
Repair based on round-00 evidence: first call rejected table alignment value Center; omit optional table alignment while preserving borders and cell padding.”
- Diagnostic and changes: first call rejected table alignment value Center; omit optional table alignment while preserving borders and cell padding.
- DSH action: failed; output D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-01\artifact.docx; render 2 pages.
- Visual evidence: D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-01\render\contact-01.png; all 2 full-page images retained.
- Repair verdict: repair did not resolve the diagnosed blocker; round-00 remains the best candidate.

## Findings
- [major] round-00 render 2 pages; source identity retained after ENGINE_FAILED (Value [Center] is invalid for the enum type [org_docx4j_wml.JcEnumeration].). Requested effect is absent; see complete contact sheet D:\开源团队作品\dsh-office-clone\the-last-docx\dist\office-aesthetic-50\cases\D20\dsh\round-00\render\contact-01.png.

## Repair history and regressions
- Round 1/5: first call rejected table alignment value Center; omit optional table alignment while preserving borders and cell padding. Result: still failed; unresolved.

## Final conclusion
**FAIL** — at least one hard gate failed or the requested effect was unavailable. Failed gates: fileIntegrity, contentPreservation, intentComplete, score_threshold. Repair was evidence-led and limited to one retry because the same verification class persisted.

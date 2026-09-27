# DOCX targeted edit comparison report

## Executive verdict

Case `docx-targeted-edit-review-v2` was completed as a first-pass comparison using the two supplied DOCX files. Both candidates passed all hard gates and exceeded the 85-point threshold. WorkBuddy scored **99/100**; the DSH DOCX plugin scored **90/100**. No repair round was used.

WorkBuddy is stronger in this case. The visible pages are almost the same, but WorkBuddy was more exact internally: the new comment covers only the quoted deployment-window sentence, while the plugin covers the entire two-sentence paragraph. WorkBuddy also keeps the executive-summary redline more granular.

The report is complete for the supplied artifacts. Runner version, model, and session logs were not supplied, so the conclusion compares these two output files rather than attributing the difference to a specific model or software revision.

## Test contract

Source: `baseline/artifact.docx`.

Frozen primary prompt:

> 请按以下已批准更新，局部修改这份报告并另存为新 DOCX，不要覆盖源文件：
>
> 1. 最新发布记录确认 Data migration 已于 2026 年 9 月 26 日收到客户签收。因此只把第 2 节 Workstream detail 表中 Workstream 为 “Data migration” 的状态从 “Awaiting client sign-off” 改为 “Client sign-off received · 26 Sep 2026”。“Acceptance testing” 行虽然状态文字完全相同，但仍在等待客户签收，必须保持不变。
> 2. 因这项签收，更新第 1 节正文统计为 “18 work items: 17 completed on schedule; 1 remains open.”；同时把 Status snapshot 表中 Completed on schedule 的 16 改为 17、Open items 的 2 改为 1。其他指标不变。
> 3. 在第 3 节 “The rollback package is verified and the transfer runbook is ready.” 这一段之后紧接着新增一句：“The transfer package was delivered to the client repository on 26 September and its SHA-256 matched the release manifest.”
> 4. 对 “The deployment window starts at 20:00 local time.” 添加一条 Word 批注，内容为 “Confirm the business owner has approved this window before go / no-go.”
>
> 上述文字与表格内容改动都要保留为 Word 修订，不要接受或拒绝修订。保留源文档中原有批注、脚注、超链接、页眉页脚、页码、表格合并与重复表头；不要改动未点名的段落、人员、日期、状态或表格行，不要重写整篇报告。版式可因新增内容自然分页，但必须保持清晰、无表格裁切、重叠或空白尾页。

Observable effects and invariants are preserved in `case.yaml`. Acceptance threshold: 85/100. Required gates: file integrity, complete render, content preservation, no critical visual defect, intent complete, targeted row only, tracked changes present, both comment anchors valid, footnote and hyperlink preserved, and repeating table header preserved.

## Environment and reproducibility

- Operating environment: Microsoft Windows NT 10.0.26200.0.
- Renderer: packaged LibreOffice 26.8.0.3 and Poppler 26.09.0.
- Render profile: 144 DPI, PNG full-page capture, same machine and installed fonts for baseline and candidates.
- Evaluation timestamp: 2026-09-26 14:52:39 +08:00.
- Structural verifier: `analyze_case02.py`; complete raw output in `structural-analysis.json`.
- Runner versions and models: not provided.

## Artifact inventory

| Artifact | SHA-256 | Pages | Evidence |
|---|---|---:|---|
| Baseline | `62591686f87390b0813b2e58eb6df520a199948122a742d08dfb9da5250bed44` | 3 | `baseline/page-1.png` through `page-3.png` |
| WorkBuddy round 00 | `750f6e9f80f7e2456820846ac71eea7951887f1b7c799fa9877a6feb002feb10` | 3 | `workbuddy/page-1.png` through `page-3.png`; `workbuddy/verifier.json` |
| Plugin round 00 | `ba69bfade0c482e0f889eaf1c816d239304ca2ac9e0b49bd9f79248e194ce143` | 3 | `plugin/page-1.png` through `page-3.png`; `plugin/verifier.json` |

## Baseline visual inventory

- Page 1: title block, executive-summary paragraphs, four-row status snapshot, review notes, and external runbook hyperlink.
- Page 2: workstream section and first ten data rows of the five-column controlled-status table.
- Page 3: repeated merged table title and column header, remaining three data rows, cutover controls, control table, closeout list, true footnote, and page footer.

All three baseline pages are readable and contain no clipping or overlap.

## WorkBuddy round 00

Action summary: updated the summary statistics and snapshot values, changed only the Data migration row, inserted the requested sentence, added the requested comment, and retained tracked changes.

Preserved regions: Acceptance testing remains unchanged; all unmentioned owners, dates, evidence, tables, headers, footers, page fields, hyperlink, existing comment, and true footnote remain intact. The workstream table still contains 13 data rows and repeats both its merged title row and column-header row across the page break.

Structural result: 12 revision nodes, two comments with valid start/end/reference anchors, one footnote reference, and the original external hyperlink. Only `word/document.xml` and `word/comments.xml` changed relative to the baseline. The new comment range contains exactly `The deployment window starts at 20:00 local time.`

Visual result: all three pages render cleanly. The longer Data migration status wraps to three lines on page 2 but remains legible and unclipped. Page 3 remains balanced and has no blank trailing page.

### WorkBuddy score

| Dimension | Score |
|---|---:|
| One-step completion | 30/30 |
| Intent-to-effect fidelity | 25/25 |
| Visual craft and aesthetic coherence | 19/20 |
| Thoroughness and delivery integrity | 15/15 |
| Modification strength and containment | 10/10 |
| **Total** | **99/100** |

## DSH DOCX plugin round 00

Action summary: completed the same visible content changes, inserted the requested sentence, added the requested comment, and retained tracked changes.

Preserved regions: Acceptance testing remains unchanged; all unmentioned owners, dates, evidence, tables, headers, footers, page fields, hyperlink, existing comment, and true footnote remain intact. The workstream table retains 13 data rows and both repeating header rows.

Structural result: 10 revision nodes, two comments with valid start/end/reference anchors, one footnote reference, and the original external hyperlink. It adds valid modern Word comment metadata parts (`commentsExtended.xml`, `commentsIds.xml`, and `people.xml`). The new comment text is correct, but its range covers both deployment-window sentences rather than only the quoted first sentence.

Visual result: all three pages render cleanly, with no clipping, overlap, broken table, or extra page. Compared with WorkBuddy, the page 1 summary revision deletes and reinserts the full statistics clause, so the visible redline is heavier than necessary.

### Plugin score

| Dimension | Score |
|---|---:|
| One-step completion | 27/30 |
| Intent-to-effect fidelity | 22/25 |
| Visual craft and aesthetic coherence | 18/20 |
| Thoroughness and delivery integrity | 15/15 |
| Modification strength and containment | 8/10 |
| **Total** | **90/100** |

## Hard-gate decisions

| Gate | WorkBuddy | Plugin |
|---|---|---|
| File integrity | PASS | PASS |
| Complete three-page render | PASS | PASS |
| Content preservation | PASS | PASS |
| No critical visual defect | PASS | PASS |
| Intent substantially complete | PASS | PASS |
| Targeted Data migration row only | PASS | PASS |
| Tracked changes present | PASS | PASS |
| Both comments structurally valid | PASS | PASS, but new range is over-broad |
| Footnote and hyperlink preserved | PASS | PASS |
| Repeating table header preserved | PASS | PASS |

## Defect register

| Severity | Runner | Location | Observed fact | Expected effect | Disposition |
|---|---|---|---|---|---|
| Major | Plugin | Page 3, structural analysis | New comment spans the whole two-sentence paragraph. | Comment should target the quoted deployment-window sentence. | Unresolved; valid comment but imprecise range. |
| Minor | Plugin | `plugin/page-1.png` | Summary redline replaces the complete statistics clause. | Minimal local revision is preferable. | Unresolved; accepted content is correct. |
| Minor | Both | Page 2 | Required longer status wraps to three lines. | Text must remain readable and unclipped. | Acceptable; no usability defect. |

No regressions or repair rounds occurred.

## Final conclusion

WorkBuddy wins this case because it is more precise, not because it looks dramatically different. Both files are usable and pass the acceptance threshold. The deciding evidence is the exact comment range and the granularity of the tracked changes. The plugin remains strong on content correctness, relationship preservation, rendering, and modern comment-part wiring, but its node-level targeting should be tightened.

This single test does not prove that WorkBuddy is universally stronger. A product-level conclusion requires additional fixed cases for creation, complex template filling, image/layout changes, damaged-document recovery, and repeated-run stability.

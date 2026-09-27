# Office DSH 200-operation rerun with visual confirmation

Date: 2026-09-27  
Repository revision under test: Office Profile from pushed code (`cce7253`), replay harness from pushed commit `ee973b5`; report added after the run.  
DSH Profile: source-built `dist/dsh-docx`, called through the DSH MCP backend.

## Verdict

PASS for the bounded backend workload: 200/200 DSH operations succeeded (50 per lane), with unique request IDs and no server stderr. This rerun also visually checked the rendered page/slide contact sheets: 34 DSH-produced DOCX-to-PDF pages, 116 PPTX slides, and 59 XLSX print-preview pages (209 individual page images total). Contact sheets and page images are retained under `dist/office-stress-20260927-visual-confirm-200/visuals/`.

The visual findings are not all clean: the HSE asbestos management plan again ends with a nearly blank page containing only the final table row, and several spreadsheets split their print areas across sparse/fractured pages. These are repeatable layout observations, not DSH extraction/read failures. Their cause (source print settings versus LibreOffice behavior) remains unresolved because Microsoft Word/Excel comparison was unavailable. No source document was repaired or modified.

This is not a visual-editing-agent benchmark: DSH parse/extract/read operations do not create edited PPTX/XLSX candidate files, so a visual quality score would be misleading. The PPTX/XLSX pages below are review-only LibreOffice exports of the inputs; the DOCX-to-PDF pages are actual DSH render outputs.

## Workload

| Lane | DSH operation | Calls | Pass | Fail | Mean latency | P95 latency |
|---|---|---:|---:|---:|---:|---:|
| DOCX | `docx-parse.execute` | 50 | 50 | 0 | 218 ms | 276 ms |
| PDF output | `docx-render.execute` | 50 | 50 | 0 | 10,305 ms | 11,496 ms |
| PPTX | `pptx.extract` | 50 | 50 | 0 | 536 ms | 610 ms |
| XLSX | `xlsx.readRange` (`A1:J20`) | 50 | 50 | 0 | 210 ms | 328 ms |
| **Total** |  | **200** | **200** | **0** |  |  |

The 200 operations were run sequentially over five office materials per lane, ten rounds each. DSH recorded 231 MCP calls including 15 artifact imports, 15 preflights, and one Doctor call. Runtime was 572.565 seconds (9m33s). No concurrency or broad semantic-correctness claim is made.

## Visual confirmation

### DSH DOCX-to-PDF output

The DSH `docx-render.execute` output was rendered at 120 DPI. All 50 render records have the same page counts and page/thumbnail PNG SHA-256 values as the prior post-push run for the matching source and round. First-versus-tenth render page images are pixel-identical for all five documents (34 pages total; zero pixels over the difference threshold and mean absolute difference 0).

| Input | Pages | Visual review |
|---|---:|---|
| `pipl-law.docx` | 10 | Dense legal text pages; orderly flow, no obvious clipping in the contact sheets. |
| `flexible-working-form.docx` | 3 | Form fields and tables visible; no obvious overflow. |
| `mhra-applicant-response.docx` | 10 | Multi-page form; later pages are sparse by design, no obvious clipping at contact-sheet scale. |
| `asbestos-management-plan.docx` | 9 | Page 9 contains only the final “removal works” table row near the top, with almost the entire page blank. Reproduced in all ten DSH renders. |
| `qa-evidence-report-template.docx` | 2 | Tables and form fields visible; no obvious overflow. |

### PPTX and XLSX review-only print previews

The five input decks and five input spreadsheets were exported with bundled LibreOffice 26.8.0.3 and rasterized through Poppler, scaled to a 480 px long edge. This review path is explicitly separate from DSH calls; no source files were edited.

| Format | Input | Pages/slides | Visual review |
|---|---|---:|---|
| PPTX | `civil-service-line-management.pptx` | 21 | Complete deck rendered; no obvious clipping in contact sheets. |
| PPTX | `civil-society-covenant.pptx` | 17 | Complete deck rendered; no obvious clipping. |
| PPTX | `prevent-duty-leadership.pptx` | 16 | Complete deck rendered; no obvious clipping. |
| PPTX | `qualifications-reform.pptx` | 28 | Complete deck rendered; tables/text visible, no obvious clipping. |
| PPTX | `timms-workshop.pptx` | 34 | Complete deck rendered; several yellow-highlighted “[Add details…]” strings are source placeholders, not conversion damage. |
| XLSX | `condition-survey-template.xlsx` | 24 | Late pages 20–24 show table/header fragments, including small isolated row groups. |
| XLSX | `green-book-appraisal-tables.xlsx` | 4 | Tables appear intact across four preview pages. |
| XLSX | `green-book-discount-factors.xlsx` | 7 | Tabular series split over seven pages; no DSH read failure implied. |
| XLSX | `qa-assumptions-log.xlsx` | 5 | Larger tables break into partial/sparse sections in the print preview. |
| XLSX | `qa-modelling-template.xlsx` | 19 | Many pages show narrow header/range fragments or sparse sheet regions. |

All 10 review-only conversions succeeded and every page image referenced by the render manifest exists (175 pages/slides; zero missing images). Every page/slide appears in the retained contact sheets. The severe spreadsheet fragmentation is a print-layout observation only; the DSH `xlsx.readRange` operation returned successfully in all 50 calls.

## Integrity, evidence, and limits

- All 15 isolated workspace files match their declared `workspaceSha256`, and those hashes remained unchanged during the run. Three DOCX inputs were intentionally sanitized before testing to remove external relationships: flexible working form (2), MHRA response (1), and asbestos plan (3). For these three, the post-sanitization workspace hash is expected to differ from the original source hash; the manifest records both. The other 12 source/workspace hashes match.
- All 200 request IDs are unique. All 50 DSH PDF-output page/thumbnail image hashes match the corresponding page images from the earlier post-push run exactly.
- This run's machine-readable DSH report is `dist/office-stress-20260927-visual-confirm-200/report.json`; fixture and operation manifest: `.../manifest.json`.
- DOCX page contact sheets and first-versus-tenth comparison data: `dist/office-stress-20260927-visual-confirm-200/visuals/visual-summary.json` and the sibling PNGs.
- Full-resolution DSH output evidence for the orphan row: `dist/office-stress-20260927-visual-confirm-200/data/objects/c9e1fbd8293146473558687e64cabf39da626fef839d6d7909bc47ade605d54e/artifact.png` (SHA-256 `c9e1fbd8293146473558687e64cabf39da626fef839d6d7909bc47ade605d54e`).
- PPTX/XLSX page images and contact sheets: `dist/office-stress-20260927-visual-confirm-200/visuals/pptx-xlsx-pages/`; case/page counts and renderer details: `render-summary.json`.
- Microsoft Word/Excel native comparison was not run. Therefore, do not attribute the observed pagination to the source file or LibreOffice alone, and do not infer equivalence with Office-native rendering.
- Native PDF input parsing is distinct from this run's DOCX-to-PDF output lane; it was exercised separately in the prior B2 `pdf-office` replay.

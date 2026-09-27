# Office stress program — engineering batches and evidence

Date: 2026-09-27
Workspace: `D:\开源团队作品\dsh-office-clone\the-last-docx`
Scope: close the implementation/diagnostic batches raised by the frozen office-stress report. The baseline, its fixtures, and output artifacts remain unchanged.

## Executive result

- B1 PPTX bridge: fixed Windows non-ASCII JSON corruption at the Python stdout boundary. Regression reproduced the default-locale failure and verified the explicit UTF-8 path; the source-built DSH MCP replay passed 50/50.
- B2 native PDF: added a separate read-only `pdf-office` DSH module using PDF.js, without routing PDFs through the DOCX renderer. Native PDF backend replay passed 50/50.
- B3 DOCX-to-PDF: measured LibreOffice as the dominant latency cost and reused one isolated LibreOffice profile for serialized calls. DSH replay passed 50/50 with exact 120-DPI page/thumbnail pixel hashes matching B0.
- B4 pagination: retained and inspected the HSE and XLSX LibreOffice print-preview findings. Word/Excel are not installed here, so renderer/source attribution is `not_run` and unresolved; no source specimen was changed or repaired.
- B5 visual repairs: added row-level DOCX table pagination control and a package-scoped XLSX print-layout action. HSE row split fixed (9 pages retained); condition survey 24→16 pages; QA model 19→9 pages (its lookup tail remains); assumptions log 5→4 pages (Welcome unchanged, wide input form 3→2 horizontal fragments with dense continuation). All source specimens remain immutable; limitations are recorded per batch.
- Final checks after B3: root typecheck passed; root tests passed (32/32); all module gates passed (504 passed, 2 skipped: optional deep-engine and environment-gated renderer tests; XLSX is 3/3); source-built DSH package build passed; provenance passed (110 module sources, zero unapproved drift). The opt-in docx-render real-runtime integration test passed during B3; its normal module-gate invocation is skipped when fixture/runtime variables are absent.

## Baseline and process

B0 is frozen at `dist/office-stress-20260927-run10`; do not edit that report, manifest, source corpus, or artifacts. Its initial report found DOCX parse 50/50, DOCX-to-PDF output rendering 50/50, PPTX extraction 20/50 (30 GBK/encoding failures), and XLSX bounded reads 50/50. Native PDF input was intentionally unsupported at that time and correctly rejected by the DOCX renderer. The 200 planned format/material operations caused 231 actual DSH calls (including associated inspection/extraction steps).

The independent integrity gate was rerun after all engineering batches: `dist/office-stress-20260927-run10/b0-baseline-verification.json` — pass, 15 specimens, 50 calls each for DOCX/PDF/PPTX/XLSX, 231 total calls, no source/workspace integrity errors. The B0 PDF cases are DOCX-rendered PDFs, not an independently sourced native-PDF corpus; B2 exercised native-PDF reading on those immutable derived PDF artifacts. The original 30 PPTX failures are preserved as historical baseline evidence; the fix is evaluated in B1, not by rewriting B0.

Each batch used an attention lock: fixed task scope, evidence first, then only the named engineering change; fixtures and unrelated modules were kept out of scope.

## Batch results

| Batch | Engineering question / change | Acceptance and evidence |
|---|---|---|
| B0 | Establish the frozen office corpus and DSH backend baseline. | `benchmarks/office-stress-b0-baseline.md`; full historical report `dist/office-stress-20260927-run10/final-report.md`; integrity check above. |
| B1 | Why did PPTX extraction fail on Windows with Unicode? Explicit UTF-8 Python bridge encoding and a GBK regression test. | 30/50 original failures reproduced under inherited encoding; fixed source-built DSH replay 50/50 with stable source/workspace hashes. Changes in `modules/pptx-office/src/index.ts` and `modules/pptx-office/tests/module.test.ts`. |
| B2 | Native PDF read versus DOCX-to-PDF output are separate capabilities. Added bounded, read-only PDF.js native inspect/text/geometry/verify in `modules/pdf-office`; retained explicit limits (no OCR, semantic table extraction, reliable reading order, PDF editing, or visual sign-off). | Work order `benchmarks/office-stress-b2-native-pdf.md`; MCP evidence `dist/office-stress-20260927-b2-native-pdf/native-pdf-50.json` (50/50; five specimens × 10; hashStable true); module tests 3/3, Profile/MCP tests 5/5. PDF.js license/route recorded in `THIRD_PARTY.md` and `OFFICE_ENGINE_DECISIONS.md`. |
| B3 | Is DOCX→PDF latency caused by rasterization? Profiled stages; LibreOffice conversion took 24.2–25.0 s, versus 0.7–2.0 s for pages and 0.15–0.37 s for thumbnails. Reused one isolated LO profile serially; preserved output contract/120 DPI. | Work order `benchmarks/office-stress-b3-render-throughput.md`; phase profile `dist/office-stress-20260927-b3-render-phase-profile.json`; warm reuse probe `dist/office-stress-20260927-b3-soffice-profile-reuse.json` (second conversion 15.904 s faster, pixels identical); DSH replay `dist/office-stress-20260927-b3-render-replay-final/docx-render-50.json` (50/50, exact page+thumbnail hashes, hashStable true; mean 10.537 s/call across five groups). Serialized reuse trades parallel throughput for reliable profile ownership; on failure the profile is retired/rebuilt. |
| B4 | Determine whether sparse/orphaned pagination is source or renderer behavior. Inspected the retained LibreOffice page evidence and checked for native Word/Excel. | Work order `benchmarks/office-stress-b4-pagination.md`; HSE asbestos plan is 9 pages, with page 9 containing only final row “removal works”; condition survey is a 24-page print preview and QA workbook a 19-page print preview with fragmented/sparse sheet regions. Word/Excel unavailable, so native comparison `not_run` and attribution unresolved. No repair attempted. |
| B5 | Repair the reproduced visual pagination issues with explicit, bounded DSH controls. DOCX: opt-in table-row `w:cantSplit`. XLSX: package-scoped worksheet print settings, avoiding ExcelJS whole-workbook rewrites. | DOCX report `benchmarks/office-stress-repair-b1-hse.md`; XLSX reports `benchmarks/office-stress-repair-b2-xlsx.md` and `benchmarks/office-stress-repair-b3-xlsx-assumptions.md`. HSE final row now stays together (9 pages before/after; full-page visual confirmation). Condition survey 24→16; QA model 19→9; assumptions log 5→4 (Welcome unchanged; input form 3→2 horizontal fragments). XLSX non-target package parts remain byte-identical and target sheet content is checked after normalizing only print settings. Dense horizontal continuation and QA lookup tail page remain explicitly reported rather than hidden. |

## B4 retained visual evidence and checksums

Renderer was the bundled LibreOffice 26.8.0.3 plus Poppler review-only export. Native Office comparison was not available; the following are observations under this renderer only.

| Evidence | SHA-256 |
|---|---|
| `dist/office-stress-20260927-run10/visuals/asbestos-management-plan-pages-03.png` | `74555b2180b687c9a101c17006b1146b55114031a8aa5bc5dc5ec73957a0dd80` |
| `.../visuals/pptx-xlsx-pages/xlsx-condition-survey-template-contact-01.png` | `2cfd1ce1ee750ceda49ec4a737c04bb75265c87b2b21edd72cdcda4ef7f44550` |
| `.../visuals/pptx-xlsx-pages/xlsx-condition-survey-template-contact-02.png` | `8f8d2fcfb6f454f0c187f125b7fd3c61ddbec4b121b4c9e70200cccc01d1741dc` |
| `.../visuals/pptx-xlsx-pages/xlsx-qa-modelling-template-contact-01.png` | `1441d7a29e07e01b7a0a92044e2791bc5a8a78673b7fcb0f70cc2dffbcf1244f` |
| `.../visuals/pptx-xlsx-pages/xlsx-qa-modelling-template-contact-02.png` | `b123171ecc70db3b7811b64953dc3bf662fd0999241b4b4cdc573524608e093c` |

The visual contact sheets and individual spreadsheet page PNGs remain in the frozen baseline visual directory. The B0 visual summary recorded 9 HSE pages and stable page hashes across its 10 repeated renders. The B4 finding is not a claim that an exported workbook should print as one page; it is a record of the observed print layout and fragments.

## Final verification

- `npm run typecheck` — pass.
- `npm test` — 4 test files, 32/32 pass.
- `npm run test:modules` — all module test outputs passed, including `pptx-office` 2/2, `xlsx-office` 2/2, and `pdf-office` 3/3; one unrelated optional deep-engine test and the environment-gated docx-render integration test were skipped by default. The real-runtime docx-render integration test was separately enabled and passed in B3.
- `npm run build:dsh` — pass; generated local package at `dist/dsh-docx`.
- `npm install --package-lock-only --ignore-scripts` — pass; dependency lock updated for `fflate` and `@xmldom/xmldom` in `xlsx-office`.
- `npm run provenance` — pass; 110 module source files hashed, zero unapproved drift. The four intentional DSH integration files are individually allowlisted in `scripts/provenance.mjs` (three B1 `docx-edit` files and the existing B3 `docx-render` adapter).
- Both 50-call DSH replays passed; final frozen B0 verification passed.

This is local source/Profile engineering evidence. It does not claim publication, remote upload, installation into a separate application, or native Microsoft Office equivalence.

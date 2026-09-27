# B4 Work Order: Pagination Findings and Renderer Attribution

## Attention lock

Investigate only the B0 visual findings: (a) the HSE asbestos-management DOCX has a nearly blank last page containing one final table row; (b) LibreOffice print previews of the Green Book condition-survey XLSX and DESNZ QA workbook span 24 and 19 pages respectively. Determine whether these are source pagination/print settings or renderer differences.

Do not edit the frozen DOCX/XLSX specimens, change `docx-render`/`xlsx-office`, or treat spreadsheet page count as an XLSX read/write failure. Keep PDF/pptx encoding and throughput batches separate.

## Evidence route

- Reuse the frozen baseline document and its retained page images; preserve their hashes.
- Check whether Microsoft Word and Excel are available for a read-only native-render comparison. If available, export copies with a recorded Office version/settings and render every resulting page for visual inspection. If unavailable, record `not_run` and leave attribution unresolved; do not infer native-app equivalence.
- Inspect the affected final/transition pages at full resolution and record source-page images, page count, renderer, font environment, and checksum.

## Acceptance

- Report the HSE orphan-row observation with exact image evidence and a renderer-attribution verdict.
- Report the XLSX page counts as print-preview observations, separate from module read/write results.
- Source and frozen baseline hashes remain unchanged.
- No document repair is attempted in this diagnostic batch; a source-specific edit requires a separate explicit task.

## Outcome

- B0 verification rerun passed: 15 specimens; 50 calls each for DOCX, PDF, PPTX and XLSX; no baseline/source integrity errors. Evidence: `dist/office-stress-20260927-run10/b0-baseline-verification.json`.
- Environment check found no `WINWORD.EXE` or `EXCEL.EXE` command and none at the standard Office installation paths. Therefore the native Office side-by-side comparison is `not_run`; attribution remains explicitly unresolved (not evidence that LibreOffice matches Office).
- Renderer evidence: bundled LibreOffice 26.8.0.3 + Poppler review-only exports, 480-px-long-edge contact sheets. `asbestos-management-plan.docx` has 9 pages; page 9 contains only the final table row “removal works,” with almost the entire page blank. This is a reproducible LibreOffice observation, not yet attributable to the DOCX pagination rules versus renderer behavior. See `dist/office-stress-20260927-run10/visuals/asbestos-management-plan-pages-03.png` (SHA-256 `74555b2180b687c9a101c17006b1146b55114031a8aa5bc5dc5ec73957a0dd80`); it also recurred in all ten B0 DSH renders.
- `condition-survey-template.xlsx` produced 24 LibreOffice print-preview pages; the contact sheets show several late pages consisting of only a small table fragment/row group (notably pages 20–24). `qa-modelling-template.xlsx` produced 19 pages; numerous pages contain only narrow header/range fragments or sparse sheet regions. These are print-layout observations only and are not XLSX extraction/read failures. All 43 page PNGs were retained; contact sheets: `xlsx-condition-survey-template-contact-01.png`, `...-02.png`, `xlsx-qa-modelling-template-contact-01.png`, `...-02.png` under `dist/office-stress-20260927-run10/visuals/pptx-xlsx-pages/`. Their SHA-256 checksums are recorded in the engineering progress report.
- No specimens or settings were modified. Frozen baseline integrity verification passed. No document repair was attempted. A native-app attribution finding still requires a host with licensed/installed Microsoft Word and Excel or a separately supplied native-render capture.

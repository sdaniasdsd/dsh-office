# B3 Work Order: DOCX-to-PDF Render Throughput

## Attention lock

Measure the three stages already used by `docx-render`: LibreOffice DOCX→PDF, Poppler full-resolution page PNGs, and Poppler thumbnails. Select one representative short and one long baseline document first. Identify where wall time is spent before changing code.

Keep the default 120 DPI, page limit, PDF output, PNG output, thumbnails, and visual-review state unchanged. Do not change PDF.js/native-PDF parsing, PPTX/XLSX modules, PDF layout/content, or the frozen baseline corpus. Do not report a lower-quality configuration as a performance fix.

## Baseline evidence

The B0 DSH run rendered five PDF-output specimens ten times each. Per-document means ranged from 33.7 s (3 pages) to 38.5 s (9 pages), with no failures; page-image SHA-256 remained identical across ten rounds per source. This batch determines whether the cost is process startup, duplicated rasterization, or document-size-dependent work, then applies only a change that preserves output quality/contract.

## Acceptance

- Record phase timings and verify the source DOCX hashes are unchanged.
- If implementing an optimization, first add a regression test for the affected stage and compare 120-DPI page pixels and thumbnail pixel dimensions/quality against the current baseline.
- Run module/Profile/type tests and a DSH MCP smoke/replay sized to the exact changed path.
- If no safe optimization preserves the declared output, document the measured bottleneck and the quality/compatibility trade-off instead of weakening defaults.

## Outcome

- Phase profile evidence: `dist/office-stress-20260927-b3-render-phase-profile.json`. On the 3-page flexible-working form, LibreOffice conversion took 24,233 ms of 25,087 ms total; Poppler pages took 708 ms and thumbnails 146 ms. On the 9-page asbestos plan, LibreOffice took 24,980 ms of 27,302 ms; Poppler pages took 1,951 ms and thumbnails 371 ms. Thus the dominant cost is LibreOffice startup/conversion, not rasterization.
- Reusing one isolated LibreOffice user profile serially cut the second 3-page conversion from 24,343 ms to 8,439 ms (15,904 ms faster); all 120-DPI page pixels matched exactly. Evidence: `dist/office-stress-20260927-b3-soffice-profile-reuse.json`.
- Implemented serialized reuse in `modules/docx-render/src/engine/adapter.ts`. On conversion failure the profile is retired and a fresh isolated profile is created for the next request; timeout now kills and awaits the child process. Output contract/default quality remains unchanged. Integration regression: `modules/docx-render/tests/profile-reuse.integration.test.ts`.
- Real DSH MCP replay: 50/50 calls passed over five DOCX specimens × 10 rounds; all 120-DPI PNG and thumbnail hashes matched the B0 reference, all source/workspace hashes remained stable. Evidence: `dist/office-stress-20260927-b3-render-replay-final/docx-render-50.json`. Mean call time was 10.54 s overall (versus B0 document means of 33.7–38.5 s); first cold run remains slower (28.05 s for pipl-law), and warm per-document means were 8.91–11.16 s.
- Verification completed: docx-render integration tests, root typecheck/tests, DSH build, and the 50-call replay. This is a single-profile serialized design: it improves repeated-call latency but intentionally does not increase concurrent LibreOffice throughput.

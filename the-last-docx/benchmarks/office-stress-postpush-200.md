# Post-push Office DSH backend regression — 200 operations

Date: 2026-09-27
Repository revision under test: `ee973b5` (test harness); Office Profile implementation: source-built `dist/dsh-docx` from the pushed Office module changes.
Runner: DSH Profile MCP stdio server; no separate installed Profile was used.

## Verdict

PASS. The four frozen office lanes each executed five materials ten times, for exactly 200 operation records. All 200 operations succeeded. The 15 source/workspace specimen hash pairs remained unchanged. All 50 DOCX-to-PDF renders reproduced the B0 page and thumbnail SHA-256 hashes exactly.

This is a backend regression/stress report, not a visual-editing-agent scorecard: none of the documents was modified, so there is no edit prompt, repair loop, or visual-agent score to invent. PDF in this 200-operation run means DOCX-to-PDF render output. Native PDF input parsing was previously exercised separately by the B2 50-call `pdf-office` replay.

## Contract and scope

- Five frozen source specimens for each lane: DOCX, PPTX, XLSX, and five B0-derived PDF render cases.
- Ten sequential DSH calls per specimen.
- Exact operations: `docx-parse.execute`; `docx-render.execute` (PDF output); `pptx-office` extract; `xlsx-office` readRange (`A1:J20` on the first worksheet returned by the frozen DSH preflight).
- Read-only operations against frozen source inputs. No fixture, baseline artifact, or generated document was manually repaired.
- Acceptance: 50/50 successful calls per lane, no missing request IDs, no source/workspace hash drift, and stable PDF-output page/thumbnail hashes against the matching B0 source's round 1.

## Environment

- DSH runner: source-built Profile at `dist/dsh-docx`, invoked through its MCP stdio server by the desktop-bundled Node runtime.
- PPTX bridge inherited Windows encoding variables were cleared by the replay runner, so the fixed explicit UTF-8 behavior was exercised without relying on a developer shell override.
- LibreOffice/Poppler runtime came from the built Profile bundle; default DOCX render settings retained.
- Sequential run duration: 591.947 seconds (9m 52s). No concurrency or semantic-correctness claim is made by this run.

## Results

| Lane | Calls | Pass | Fail | Mean latency | P95 latency |
|---|---:|---:|---:|---:|---:|
| DOCX parse | 50 | 50 | 0 | 222 ms | 275 ms |
| DOCX → PDF render | 50 | 50 | 0 | 10,624 ms | 12,037 ms |
| PPTX extract | 50 | 50 | 0 | 538 ms | 611 ms |
| XLSX `A1:J20` read | 50 | 50 | 0 | 223 ms | 352 ms |
| **Total** | **200** | **200** | **0** | — | — |

The DSH server handled 231 MCP tool calls in total: 200 workload operations, 15 artifact imports, 15 structure preflights, and one Doctor call. All 200 workload records have unique request IDs. The server stderr stream was empty.

## Integrity and render comparison

- All 15 source and isolated-workspace SHA-256 values after the run match the values recorded in the run manifest.
- For every PDF-output source and round (50 total), `pageCount`, each full-page PNG SHA-256, and each thumbnail SHA-256 match that source's B0 round-1 result exactly.
- The 200 operation records, request IDs, durations, summaries and per-call errors (empty on success) are retained in `dist/office-stress-20260927-postpush-200/report.json`.
- The isolated manifest, material sources, workspace copies, and DSH artifact data are under `dist/office-stress-20260927-postpush-200/`; the runner never targeted or wrote into the frozen B0 run.

## Follow-up and limits

This validates successful bounded calls and repeatable render output for these five official-source samples per format. It does not assert broad semantic correctness of extracted content, stress concurrency, or compare LibreOffice rendering against Microsoft Word/Excel. The latter Office-native pagination attribution remains unavailable in this environment. Native PDF input parse evidence remains in `dist/office-stress-20260927-b2-native-pdf/native-pdf-50.json`.

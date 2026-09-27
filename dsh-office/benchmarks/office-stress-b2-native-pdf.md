# B2 Work Order: Native PDF Input Capability

## Attention lock

Only add and validate read-only native PDF inspection/text extraction in the DSH Office Profile. Reuse the already selected PDF.js route from `D:\开源团队作品\pdf分区`; keep PDF input distinct from the existing `docx-render` DOCX-to-PDF output path.

Do not change DOCX rendering, its latency/DPI behavior, PPTX/XLSX modules, PDF editing/creation/OCR, the frozen B0 corpus, or the B1 encoding fix in this batch.

## Why this batch exists

The B0 native-PDF boundary probe proved that `docx-render.inspect` correctly rejects an existing PDF as `FORMAT_MISMATCH`. That is a capability boundary, not a defect in DOCX rendering. The product workspace already contains PDF.js-backed native PDF parsing modules, and this DSH profile needs a separate PDF contract/tool to expose that capability safely.

## Implementation route

- Use Mozilla PDF.js (`pdfjs-dist`), Apache-2.0, the parser selected and already used by the sibling PDF workspace.
- Start with bounded, read-only inspect + native text/geometry extraction and verify; do not claim OCR, table semantics, reliable reading order, visual approval, or editing.
- Preserve artifact immutability, SHA-256 source identity, page/node/input/output budgets, cancellation/timeout handling, and explicit encrypted-PDF behavior.
- Do not route PDF inputs through `docx-render` or weaken that module's format check.

## Acceptance

- Focused unit and profile/MCP tests cover a valid native-text PDF, non-PDF rejection, malformed PDF, encrypted PDF behavior, page/input/node limits, and source hash binding.
- Through the DSH MCP stdio backend, parse the five already frozen PDF stress specimens ten times each (50 calls); retain per-call request ID, duration and result/error.
- Check source specimen hashes before and after; no source mutation.
- Build the DSH Profile and run focused + repository tests/typecheck appropriate to the changed surface.
- Update the master report with capability boundary before and after, exact coverage limits and evidence paths.

## Status

- [x] Design and implementation (`modules/pdf-office`, PDF.js 6.3.289)
- [x] Focused tests and profile integration (module 3/3, Profile + MCP 5/5, TypeScript and DSH package build)
- [x] DSH backend 5 x 10 replay (50/50 pass; isolated replay store; frozen source/workspace/PDF hashes stable)
- [x] Master report updated

# pptx-office

DSH adapter for [Python `python-pptx` 1.0.2](https://python-pptx.readthedocs.io/en/latest/) (MIT). It inspects `.pptx` shape/text/run structure, performs guarded edits to explicitly addressed text runs, applies a bounded all-slides title/body typography hierarchy, and provides allow-listed cover design profiles; it is not a presentation renderer.

Install the pinned runtime dependency into the Python interpreter configured by the Profile with `python -m pip install -r modules/pptx-office/requirements.txt`. The module never installs packages itself. Macro-enabled `.pptm` packages are rejected; `replaceText` requires a source-revision-specific shape/run address and exact `expectedText` match. `formatText` accepts one `{ scope: "allSlides", titleFontSize, bodyFontSize, accentColor }` change, sets title/center-title placeholder sizes only when the requested size fits the measured placeholder box, and only reduces explicitly oversized runs in body/subtitle placeholders (it never enlarges body text or reformats arbitrary diagram copy). It applies the accent color to title placeholders only when contrast against a readable solid slide background is at least 4.5:1. If the requested color would be low-contrast, black or white is selected; on a complex or unreadable background the original title color is preserved. It does not change slide masters, theme fonts, charts, tables, or non-placeholder text. Results are immutable artifacts; rendering and visual review are still required before delivery.

Through MCP, import the deck with `docx_import`, then call `pptx_call` with `operation: "execute"` and `input: { requestId, artifactRef, payload: { action: "extract" } }`. Each replacement entry is `{ slideNumber, shapeId, paragraphIndex, runIndex, expectedText, replaceWith }` under `payload.changes`. Use the extracted indices only with the same source artifact revision. For restrained deck-wide hierarchy, call with `payload: { action: "formatText", changes: [{ scope: "allSlides", titleFontSize: 30, bodyFontSize: 18, accentColor: "#244A67" }] }`.

## Cover design system

The design system takes architectural inspiration from [PPT Master](https://github.com/hugohe3/ppt-master) (MIT): style decisions live in a versioned catalog outside the generic editing path; the output remains native and editable; a saved/reopened artifact still requires rendered visual review. This module has no runtime dependency on PPT Master and does not embed its assets or generation workflow.

Two cover-only profiles are available:

- `indigo-paperlight-v1`: dark editorial cover with a packaged paper-light bitmap, contrast-safe title treatment, and editable native art word.
- `cool-corporate-field-v1`: an entirely native, editable editorial composition. It reserves the source title's reading space, adds a deep-navy visual field and one gold focal dot, then adds a single editable art word. Its three-color palette is intentionally restrained so the source title remains the visual priority.

Invoke either with `payload: { action: "applyArtStyle", styleId: "cool-corporate-field-v1", artWord: "DSH" }`.

Profiles deliberately decorate the first slide only. They do not edit masters, themes, layouts, charts, tables, later slides, or source text. New profiles require an allow-listed `styleId`, save/reopen text-preservation coverage, and rendered visual review. This keeps the generic engine stable while allowing the application layer to grow a style catalog safely.

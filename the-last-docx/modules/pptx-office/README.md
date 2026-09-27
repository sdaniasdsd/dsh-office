# pptx-office

DSH adapter for [Python `python-pptx` 1.0.2](https://python-pptx.readthedocs.io/en/latest/) (MIT). It inspects `.pptx` shape/text/run structure and performs guarded edits to explicitly addressed text runs; it is not a presentation renderer.

Install the pinned runtime dependency into the Python interpreter configured by the Profile with `python -m pip install -r modules/pptx-office/requirements.txt`. The module never installs packages itself. Macro-enabled `.pptm` packages are rejected; `replaceText` requires a source-revision-specific shape/run address and exact `expectedText` match. Results are immutable artifacts, and visual layout review is still required before delivery.

Through MCP, import the deck with `docx_import`, then call `pptx_call` with `operation: "execute"` and `input: { requestId, artifactRef, payload: { action: "extract" } }`. Each replacement entry is `{ slideNumber, shapeId, paragraphIndex, runIndex, expectedText, replaceWith }` under `payload.changes`. Use the extracted indices only with the same source artifact revision.

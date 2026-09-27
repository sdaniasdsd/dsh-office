# xlsx-office

DSH adapter for [ExcelJS 4.4.0](https://github.com/exceljs/exceljs) (MIT). It inspects `.xlsx` workbooks, reads bounded cell ranges, writes literal cell values to a new immutable artifact, creates workbooks, and verifies reopened output. Macro-enabled `.xlsm` is intentionally unsupported.

ExcelJS does not calculate formulas. Existing formula text and any cached result are reported separately; formula caches can be stale. Cell writes accept only literal strings, numbers, booleans, or null, and never evaluate user-provided formulas. Run `npm install` at the workspace root to install ExcelJS.

The ZIP preflight reports feature families ExcelJS may not preserve (for example charts/drawings, external links, pivots, embedded objects, ActiveX controls, signatures, and selected newer metadata) and refuses cell edits for those workbooks. The `setPrintLayout` action is different: it patches only each explicitly named worksheet's `sheetPr/pageSetUpPr` and `pageSetup` XML and leaves all other package parts byte-for-byte unchanged. Signed workbooks are refused because any package edit invalidates the signature. It does not change print areas, cell data, styles, formula caches, or defined names.

The engine has structural safeguards, but visual layout still needs Excel or LibreOffice review when appearance matters. Workbook output is not a visual-rendering guarantee.

Through MCP, import a workbook with `docx_import`, then call `xlsx_call` with `operation: "execute"` and `input: { requestId, artifactRef, payload: { action: "readRange", sheet: "Data", range: "A1:D20" } }`. For edits, use `payload: { action: "setCells", changes: [{ sheet: "Data", address: "B2", value: 42 }] }`; output is a separate artifact.

For a print-only copy use `payload: { action: "setPrintLayout", sheets: [{ sheet: "Data", orientation: "landscape", fitToWidth: 1, fitToHeight: 0 }] }`. Fit dimensions are explicit per sheet; omitted width/height default to one page wide and unlimited pages tall. This operation produces a separate artifact and returns `visualReview: "pending"`; render and inspect every page before treating the layout as accepted.

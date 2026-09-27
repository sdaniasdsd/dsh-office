# xlsx-office

DSH adapter for [ExcelJS 4.4.0](https://github.com/exceljs/exceljs) (MIT). It inspects `.xlsx` workbooks, reads bounded cell ranges, writes literal cell values to a new immutable artifact, creates workbooks, and verifies reopened output. Macro-enabled `.xlsm` is intentionally unsupported.

ExcelJS does not calculate formulas. Existing formula text and any cached result are reported separately; formula caches can be stale. Cell writes accept only literal strings, numbers, booleans, or null, and never evaluate user-provided formulas. Run `npm install` at the workspace root to install ExcelJS.

The ZIP preflight reports feature families ExcelJS may not preserve (for example charts/drawings, external links, pivots, embedded objects, ActiveX controls, signatures, and selected newer metadata) and refuses cell edits for those workbooks.

The engine has structural safeguards, but visual layout still needs Excel or LibreOffice review when appearance matters. Workbook output is not a visual-rendering guarantee.

Through MCP, import a workbook with `docx_import`, then call `xlsx_call` with `operation: "execute"` and `input: { requestId, artifactRef, payload: { action: "readRange", sheet: "Data", range: "A1:D20" } }`. For edits, use `payload: { action: "setCells", changes: [{ sheet: "Data", address: "B2", value: 42 }] }`; output is a separate artifact.

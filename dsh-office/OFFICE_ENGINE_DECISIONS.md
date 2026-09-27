# Office engine decisions: PPTX, XLSX, and native PDF

These choices adapt two established open-source engines to the existing DSH module contract (`inspect` / `execute` / `verify`), managed immutable artifacts, and bounded inputs. They do not replace the existing DOCX modules.

## PPTX — python-pptx 1.0.2

Selected: [python-pptx](https://github.com/scanny/python-pptx), [documentation](https://python-pptx.readthedocs.io/en/latest/), MIT.

It offers a high-level Python object model for opening and saving presentations, slides, shapes, text frames, runs, and tables, which maps directly to DSH's inspect/extract/precise-run-edit workflow. DSH calls it through a small JSON subprocess bridge so the TS host retains artifact ownership and resource limits. Its scope is editing/structure—not rendering. `.pptm` is rejected; run replacement requires a stable-in-this-source-revision shape/run address and exact precondition text. Visual review remains separate.

The engine is pinned in `modules/pptx-office/requirements.txt`. It is never installed automatically. Install it into the interpreter configured in the DSH Profile.

## XLSX — ExcelJS 4.4.0

Selected: [ExcelJS](https://github.com/exceljs/exceljs), MIT.

It is a mature TypeScript/Node workbook library with workbook, worksheet, cell, value, and formula-cache APIs, avoiding a second Python bridge for ordinary spreadsheet operations. DSH exposes bounded range reads, literal-value cell edits to a new immutable artifact, workbook creation, and reopen verification. Formula values are returned with their cached results; ExcelJS does not calculate formulas. Formula objects are not accepted as write input. `.xlsm` is rejected. Known feature families ExcelJS may not preserve (such as charts, drawings, external links, pivots, embedded/ActiveX parts, signatures, and selected newer metadata) are reported and block cell edits.

ExcelJS is pinned in the workspace package manifest and lockfile. Structural verification does not certify visual layout or formula freshness; use Excel/LibreOffice when those matter.

For print pagination, `xlsx-office` also offers an opt-in `setPrintLayout` action implemented as a narrow SpreadsheetML package patch with `fflate` and `@xmldom/xmldom`. It only changes `pageSetUpPr/@fitToPage` and `pageSetup` on named sheets; it does not serialize workbook objects through ExcelJS, since doing so could drop unsupported features that the cell-write path correctly refuses. Other ZIP parts are checked byte-for-byte, and target worksheet XML must be identical outside print-setup attributes. Signed workbooks are refused. Layout acceptance still requires visual rendering; the operation does not infer print areas or suppress worksheets. This follows the OOXML placement of the fit flag and page constraints ([`pageSetUpPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pagesetupproperties?view=openxml-3.0.1), [`pageSetup`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/7d79dcce-d8fb-43da-a189-f0e46658da62)).

## Native PDF read — Mozilla PDF.js 6.3.289

Selected: [Mozilla PDF.js](https://github.com/mozilla/pdf.js), distributed as [`pdfjs-dist`](https://www.npmjs.com/package/pdfjs-dist), Apache-2.0. This is the same engine already used by the sibling `D:\开源团队作品\pdf分区` implementation.

The DSH `pdf-office` extension accepts existing PDF artifacts and exposes bounded metadata inspection, native text items with PDF user-space geometry, and page-count/text verification. It does not route PDFs through `docx-render`, nor claim OCR, semantic paragraphs, reliable reading order, table structure, active-content scanning, visual approval, PDF editing, or creation. Password-protected inputs fail with an explicit `PASSWORD_REQUIRED` classification; there is no password collection or persistence flow. Input bytes, page count, text characters/items, JSON output and wall time are bounded, but PDF.js runs in-process and these limits are not an OS-level memory/CPU sandbox.

The DSH adapter depends on `pdfjs-dist` as an exact package dependency instead of inlining its worker/runtime files into the Profile bundle. Upgrade requires re-running native-PDF module tests and the 5-material × 10-call DSH replay.

PPTX, XLSX, and PDF engines intentionally keep engine-specific behavior behind separate DSH module IDs and focused MCP tools (`pptx_call`, `xlsx_call`, `pdf_call`) while remaining available through the generic module registry.

Both modules intentionally keep engine-specific behavior behind separate DSH module IDs and provide focused MCP tools (`pptx_call`, `xlsx_call`) while remaining available through the generic module registry.

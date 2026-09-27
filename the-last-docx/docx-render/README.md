# docx-render

`docx-render` makes DOCX pages visible to an Agent. Its default adapter follows the document-skill workflow: LibreOffice converts DOCX to PDF, Poppler turns the PDF into page PNGs and low-resolution thumbnails, and the `ArtifactStore` publishes those files as normal Profile artifacts.

The module does not parse or edit DOCX structure. The caller can use `docx-parse` / `docx-complex-parse` to obtain semantic IDs and source pointers, then attach those references to `VisualFinding` records after reviewing the rendered pages. This keeps visual evidence here while leaving node lookup and repairs to parsing/editing/orchestration.

## Runtime and dependencies

The default engine requires `soffice` (LibreOffice) and `pdftoppm` (Poppler) in the Agent execution environment. Their executable paths are Profile configuration; the module does not install software or assume that the user's desktop has either tool. For a remote preview service, inject another `RenderEngine` implementing the same small adapter contract.

The module uses no `Mammoth` fallback: HTML conversion does not provide rendered pages. If the configured renderer is unavailable, the stable error is `ENGINE_UNAVAILABLE`, not a successful but weaker visual-verification result.

## Usage

```ts
import { createDocxParseModule } from '@dsh-office-profile/docx-parse';

const parser = createDocxParseModule();
const renderer = createDocxRenderModule({ artifactStore });
const parsed = await parser.handlers.execute({ artifactRef, requestId: 'parse-1', operation: 'execute' });
const currentArtifact = parsed.result.artifact;
const rendered = await renderer.handlers.execute({ artifactRef: currentArtifact, requestId: 'render-1', operation: 'execute' });

// The Agent visually reviews rendered.result.pages[].image (and thumbnails).
const checked = await renderer.handlers.verify({
  artifactRef: currentArtifact,
  requestId: 'verify-1',
  operation: 'verify',
  renderResult: rendered.result,
  parseResult: parsed.result,
  visualFindings: [], // explicit empty array means reviewed and no visual defects found
});
```

If `visualFindings` is omitted and `requireVisualReview` is enabled, `verify` reports a partial result. It checks that page images exist; it does not claim to automatically detect every visual defect. Findings can include parser-provided `sourceSemanticIds` and `sourcePointers`; when they do, pass the matching `docx-parse` execute result so the module can reject stale or unknown source references. `docx-render` imports that public type and does not copy or change the parser IR contract.

The original artifact is read-only. PDF, page images, and thumbnails are written as derived artifacts via `ArtifactStore.write()`. Temporary files use a unique directory and are removed when rendering completes or fails. Structural validity remains the parser/inspector's responsibility; `inspect` here only performs a lightweight ZIP signature and declared-type preflight.

## Standalone development

```sh
npm install
npm run typecheck
npm test
```

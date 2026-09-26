# docx-edit

`docx-edit` is the structured-editing module in the DOCX Profile group. It uses `@docx4j/core-ts` as its OOXML engine and consumes node anchors from the public `docx-parse` contract.

The module owns its Profile-facing input, output, warnings, errors, and verification report. The engine is confined to `src/engine/adapter.ts`; no docx4j object is returned across the module boundary. Artifact bytes are read and written through an injected `ArtifactStore`, which is where Profile should connect `office-files`.

## Current operation shape

```ts
const parsed = await docxParse.handlers.execute({ artifactRef, operation: 'execute', requestId });
const target = targetFromDualIR(parsed.result.ir.content, semanticId);

const edited = await docxEdit.handlers.execute({
  artifactRef: parsed.result.artifact,
  operation: 'execute',
  requestId,
  plan: {
    edits: [{
      kind: 'replaceText',
      target,
      find: 'draft',
      replace: 'final',
      revision: 'track',
    }],
  },
});
```

The source document is never overwritten by the engine. `ArtifactStore.write()` receives the edited bytes and creates a new artifact. Each target is re-identified in the current DOCX using its `paraId` when available, otherwise its parser-generated content digest. Missing, stale, or ambiguous targets fail before a new artifact is written.

`replaceText` supports a literal match within one paragraph, including text split across Word runs. If a match occurs more than once, provide its zero-based `occurrence`. `revision: 'respectDocument'` leaves the document's existing tracking mode in effect; `'track'` writes this edit as a Word revision; `'untracked'` temporarily disables tracking for this operation. `addComment` can anchor to a quoted range or to the whole paragraph. `resolveRevisions` accepts or rejects the revisions in one target paragraph; Profile must explicitly enable `allowAcceptReject` before using it.

## Formatting in place, and why it is the annotation-safe path

`formatParagraph` changes the target paragraph's own properties without rebuilding the package:

```ts
const formatted = await docxEdit.handlers.execute({
  artifactRef: parsed.result.artifact,
  operation: 'execute',
  requestId: 'restyle-1',
  plan: {
    edits: [
      {
        kind: 'formatParagraph',
        target: targetFromDualIR(parsed.result.ir.content, chapterId),
        styleId: 'Heading1',
        outlineLevel: 1,
        spaceAfter: 12,
        font: { bold: true, size: 18, color: '#1F3864' },
      },
      { kind: 'formatParagraph', target: targetFromDualIR(parsed.result.ir.content, introId), spaceAfter: 8, lineSpacing: 15 },
    ],
  },
});
```

Supported properties: `styleId`, `outlineLevel`, `alignment`, `spaceBefore`, `spaceAfter`, `lineSpacing`, and a `font` object (`bold`, `italic`, `size`, `color`, `name`). `outlineLevel` follows Office JS, not the XML: **1-9 for heading levels, 10 for body text** (`w:outlineLvl` is zero-based; this field is not). Spacing and size are points; `color` is `'#RRGGBB'`. At least one property is required — a request that would change nothing fails with `INVALID_INPUT` rather than reporting a silent no-op.

The module writes `w:pPr` on that paragraph, and for `font` the direct run properties of its runs. Every other part keeps its original bytes, so `comments.xml`, `footnotes.xml`, headers, numbering and custom XML survive untouched. That is the difference that matters when a document must be re-styled for review but its annotated content must not move: the rebuilding modules (`docx-create`) construct a new package from a plan and have no comments or footnotes to carry over, so they cannot make the same promise. Formatting in place is therefore the operation to reach for whenever a document carries annotations.

Applying formatting and *verifying* it are separate steps: `execute` records what was requested in `expectedFormats`, and `verify` reopens the saved package and asserts every requested property is actually present (`format.paragraph.N` checks). A property that did not land fails verification instead of passing silently.

`styleId` requires the style to exist in the document. Assigning a style id the document does not define writes `w:pStyle` and changes no appearance, which is why `font` exists: direct character formatting produces a visible hierarchy in a document whose `styles.xml` carries no heading styles. The module does not create styles.

`allowFormatting` (default `true`) gates this operation, independently of the other edit kinds.

## Current operation scope

The operation surface is paragraph-scoped, including paragraphs inside table cells: `replaceText`, `addComment`, `resolveRevisions`, `formatParagraph`, `formatTable`, and `insertParagraph`.

`insertParagraph` adds one paragraph `Before` or `After` a parsed target:

```ts
{ kind: 'insertParagraph', target, text: 'Hiring resumed in Q3.', position: 'After', revision: 'track' }
```

With `revision: 'track'` the new paragraph and its runs carry `w:ins` revision marks, the same way `replaceText` writes a `w:ins`/`w:del` pair - the content API asks the package's change tracker for its marks, so an insertion is a revision rather than an untracked paragraph. `verify` reports `insert.N`, and it separates the two facts that matter: the text being present, and the text being present **as a revision**. `allowBlockEdits` (default `true`) gates this kind.

It still does not remove paragraphs or tables, edit headers and footers, or resolve revisions across the entire document. Those operations can be added behind the same adapter after their target and preservation rules are specified.

## External relationships are an invariant, not a gate

An in-place edit writes text and properties. It never creates, drops or retargets a relationship, so the Profile lets a document that already carries an external hyperlink through the edit pre-flight (see `DEFAULT_EXECUTE_POLICIES` in `src/profile.ts`) - refusing the whole document up front also refused the ordinary case of a project report with a runbook link that needed one number changed. Rendering and delivery stay strict, because a renderer or a recipient may resolve links on its own.

What replaces the coarse gate is a precise one: `execute` compares the package's external-relationship set before and after the write and fails with `RELATIONSHIP_INTEGRITY_FAILED` if anything was added, dropped or retargeted. The deployer can still tighten or widen the pre-flight policy per module through `ProfileOptions.safetyPolicy`.

## Engine choice

## Table instance geometry

`formatTable` writes the table element's own geometry, which is what decides the layout:

```ts
const formatted = await docxEdit.handlers.execute({
  artifactRef: styled.result.artifact,
  operation: 'execute',
  requestId: 'table-1',
  plan: { edits: [{
    kind: 'formatTable',
    target: tableTargetFromDualIR(parsed.result.ir.content, tableId),
    styleId: 'TableGrid',
    width: 451,
    layout: 'fixed',
    columnWidths: [225.5, 225.5],
    cellMargins: { top: 0, left: 5.4, bottom: 0, right: 5.4 },
    borders: { top: { size: 0.5, color: '#9AA6B2' }, insideH: { size: 0.5, color: '#9AA6B2' } },
    headerRow: true,
    cellVerticalAlignment: 'center',
  }] },
});
```

Supported: `styleId`, `width`, `layout`, `alignment`, `columnWidths`, `cellMargins`, `borders`, `headerRow`, `cellVerticalAlignment`. Units are points. When both `width` and `columnWidths` are given they must agree, because a fixed-layout grid that does not add up to the table width is not a layout Word will honour.

A table target names a body-level table (`/w:document/w:body/w:tbl[n]`) and is resolved through `tableTargetFromDualIR`; `targetFromDualIR` keeps refusing non-paragraph anchors, so a paragraph edit can never be quietly aimed at a table.

**Why this is not part of the style module.** A table style only says what a table *may* look like. A `<w:tbl>` whose `<w:tblGrid>` carries bare `<w:gridCol/>` children declares **no preferred widths at all**, so the renderer falls back to the minimum width the content needs and a cell reading `A1` wraps onto two lines. No style definition can repair that - the geometry has to be written on the table. It is also where every mature library puts it: python-docx exposes `autofit`/`column.width`/`cell.width` on the table object, Apache POI on `XWPFTable`, docx4j on the `Tbl`. docx4j-core-ts offers no writable accessor for `w:tblW`, `w:tblGrid` or `w:tblCellMar`, so `engine/table-xml.ts` edits the tree directly.

Table geometry is applied in a second pass, after the paragraph edits, because it rewrites the main document part and would otherwise read a stale copy. `verify` reads the saved table back and reports `table.N.geometry` and `table.N.grid` - a grid whose columns all read zero means the wrapping will come back, however confidently the request was written.

`allowTableFormatting` (default `true`) gates this kind independently.

## Engine choice

The dependency is pinned to `@docx4j/core-ts` 0.1.5. The package's editable model stays behind one adapter, making a later engine change independent of the Profile contract. Since this is a new upstream implementation, the module should keep growing real-document regression fixtures, especially around older revisions, comments, tables, images, fields, and extension parts.

## Standalone development

```sh
npm install
npm run typecheck
npm test
```

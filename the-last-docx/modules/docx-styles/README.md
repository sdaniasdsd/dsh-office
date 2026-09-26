# docx-styles

`docx-styles` owns one OOXML part: the style definitions. It resolves, audits and authors `styles.xml` through `@docx4j/core-ts`, and it does not touch document content — applying a style to a paragraph stays with `docx-edit`, so the two concerns do not each grow a copy of the other.

The split follows the part taxonomy every mature OOXML toolkit uses: docx4j's `StyleDefinitionsPart`, Apache POI's `XWPFStyles`, the Open XML SDK's `StyleDefinitionsPart`. Those are separate from the content-editing API for the same reason it is separate here.

No Python. The module is TypeScript over docx4j, so it adds no runtime to a deployment that already has the engine.

## Why it exists

`docx-edit` writes `w:pStyle w:val="Heading2"` on a paragraph. That is a *reference*. If the styles part never defines `Heading2`, the reference resolves only because Word happens to synthesize its built-in styles — the document does not control the appearance, and a renderer that does not synthesize them shows body text. And when the appearance lives in per-paragraph direct formatting instead of in a style, nothing keeps a document consistent.

`docx-create` does author a full `styles.xml`, but it builds a new package from a plan, so it cannot carry a document's comments and footnotes across. That left a gap: no module could put named styles into an *existing* document.

## Operations

### inspect

```ts
const inventory = await docxStyles.handlers.inspect({ artifactRef, requestId, operation: 'inspect' });
```

Returns every definition with the property groups it actually carries, the default paragraph and table styles, and the content's references split three ways:

- `defined` — the part defines the referenced id.
- `builtInUndefined` — the reference is to a Word built-in the part does not define. Word synthesizes these, so the link is not broken; it is **renderer-dependent**, which is a milder and different fact.
- `unknownUndefined` — neither defined nor built-in. This is a genuinely dangling reference.

Keeping those apart is the point: reporting every undefined built-in as "broken" would be noise, and reporting none of them would hide the reason an appearance is not reproducible.

### execute

```ts
const styled = await docxStyles.handlers.execute({
  artifactRef, requestId: 'restyle-styles', operation: 'execute',
  plan: { styles: [
    { styleId: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal',
      paragraph: { outlineLevel: 0, keepNext: true, spacing: { before: 0, after: 10 },
                   bottomBorder: { style: 'single', size: 0.5, color: '#1F3864', space: 3 } },
      run: { bold: true, size: 18, color: '#1F3864', font: { ascii: 'Arial', eastAsia: 'SimHei' } } },
  ] },
});
```

Definitions named by the plan are replaced; **every other definition in the part is kept exactly as it was**. A module that rewrote the whole part would silently drop every style the document already carried.

Supported: `styleId`, `name`, `type` (`paragraph`/`character`/`table`), `basedOn`, `next`, `link`, `default`, `quickFormat`, a `paragraph` block (`outlineLevel`, `keepNext`, `keepLines`, `alignment`, `spacing`, `indentation`, `bottomBorder`), a `run` block (`bold`, `italic`, `size`, `color`, `font`), and a `table` block (`width`, `layout`, `cellMargins`, `borders`).

**Units are points throughout**, as in Office JS and in `docx-edit`'s `formatParagraph`. The engine converts to what the format stores — twips (`pt × 20`), half-points (`pt × 2`), eighths of a point for border sizes — so a caller never does unit arithmetic and cannot silently write the wrong scale.

`paragraph.outlineLevel` is the **raw zero-based OOXML value** (0-8), not Office JS's 1-based `outlineLevel`. A style definition is the one place where the stored number is the honest one to write.

### verify

Audits a package without needing to know what was asked for, and with `expectations.styles` also checks the request landed:

| check | fails when |
|---|---|
| `styles.part.present` | the package carries no styles part |
| `styles.order.valid` | a style's children are out of the sequence ECMA-376 declares |
| `styles.definitions.present` | a requested id is absent from the saved part |
| `styles.definitions.carry` | a definition names properties but carries none |
| `styles.references.resolve` | the content references a non-built-in style nothing defines |
| `styles.references.builtin` | informational: built-ins referenced but not defined here |

`styles.order.valid` exists because element order in these parts is not cosmetic. Word offers to "repair" a file whose `w:pPr` puts `w:pBdr` after `w:spacing`, or whose `w:style` puts `w:pPr` after `w:rPr`. The emitter builds every container through an explicit ordering table, so the order is right by construction; the check catches a part that arrived out of order from somewhere else.

## Configuration

`limits`: `maxInputBytes` (64 MiB), `maxOutputBytes` (128 MiB), `maxStyles` (500). `timeoutMs`: 30000.

`featureFlags`:
- `allowStyleDefinitions` (default `true`) — gates writing definitions at all; inspect and verify keep working when it is off.
- `allowRedefineExisting` (default `true`) — rewriting a definition the document already has changes the appearance of every paragraph using it, so it is a separate decision from adding a style the document lacks.

## Composition

Styles must be defined before content references them. The Profile orders that: `docx-styles execute` first, then `docx-edit`, whose `formatParagraph` writes the references. The intermediate package then holds styles nothing uses yet, which is valid and carries no dangling reference.

## Standalone development

```sh
npm install
npm run typecheck
npm test
```

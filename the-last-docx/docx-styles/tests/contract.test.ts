import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { DEFAULT_FEATURE_FLAGS, DEFAULT_LIMITS } from '../src/config';
import type { StyleSpec } from '../src/domain/docx-styles';
import { Docx4jStylesEngine } from '../src/engine/adapter';
import { styleElement } from '../src/engine/style-xml';
import { NS } from '../src/engine/xml';

const richPath = fileURLToPath(new URL('../../docx-parse/fixtures/_generated/rich.docx', import.meta.url));

async function richFixture(): Promise<Uint8Array> {
  return new Uint8Array(await readFile(richPath));
}

/**
 * A minimal package whose content references the given style ids, with a styles
 * part that defines none of them. `createPackage` ships a default style set, so
 * the part is emptied explicitly - otherwise `Heading2` would already be defined
 * and the undefined-reference cases could not be exercised at all.
 */
async function packageReferencing(references: readonly string[]): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
  const body = references
    .map((id) => `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr><w:r><w:t>${id}</w:t></w:r></w:p>`)
    .join('');
  pkg.getMainDocumentPart().setXml(`<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  const { StyleDefinitionsPart } = await import('@docx4j/core-ts/parts');
  const main = pkg.getMainDocumentPart();
  const styles = main.styleDefinitionsPart ?? new StyleDefinitionsPart();
  if (!main.styleDefinitionsPart) main.addTargetPart(styles);
  styles.setXml(`<w:styles ${NS}/>`);
  return pkg.save();
}

const packageWithEmptyStyles = () => packageReferencing([]);

const HEADINGS: StyleSpec[] = [
  {
    styleId: 'Heading1',
    name: 'heading 1',
    basedOn: 'Normal',
    next: 'Normal',
    quickFormat: true,
    paragraph: { outlineLevel: 0, keepNext: true, spacing: { before: 0, after: 10 }, bottomBorder: { style: 'single', size: 0.5, color: '#1F3864', space: 3 } },
    run: { bold: true, size: 18, color: '#1F3864', font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'SimHei' } },
  },
  {
    styleId: 'Heading2',
    name: 'heading 2',
    basedOn: 'Normal',
    next: 'Normal',
    quickFormat: true,
    paragraph: { outlineLevel: 1, keepNext: true, spacing: { before: 16, after: 7 } },
    run: { bold: true, size: 13, color: '#2E5496', font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'SimHei' } },
  },
];

describe('docx-styles engine', () => {
  it('defines the named styles so the content’s references stop being renderer-dependent', async () => {
    const engine = new Docx4jStylesEngine();
    const result = await engine.execute(await richFixture(), HEADINGS, DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS);

    const heading1 = result.inventory.styles.find((style) => style.styleId === 'Heading1');
    const heading2 = result.inventory.styles.find((style) => style.styleId === 'Heading2');
    expect(heading1?.carries).toContain('rPr.b');
    expect(heading1?.carries).toContain('rPr.color');
    expect(heading1?.carries).toContain('pPr.pBdr');
    expect(heading2?.carries).toContain('rPr.color');
    expect(result.defined.map((entry) => entry.action)).toEqual(['replaced', 'created']);
    expect(result.defined[0]?.properties).toContain('paragraph.bottomBorder');
    expect(result.defined[1]?.properties).toContain('run.size');
    expect(result.inventory.references.unknownUndefined).toEqual([]);

    expect(await engine.verify(result.bytes, DEFAULT_LIMITS, { styles: HEADINGS })).toMatchObject({ ok: true });
  });

  it('keeps the styles the document already carried instead of rewriting the part', async () => {
    const engine = new Docx4jStylesEngine();
    const before = await engine.inspect(await richFixture(), DEFAULT_LIMITS);
    const result = await engine.execute(await richFixture(), [HEADINGS[1]!], DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS);

    const survivors = result.inventory.styles.map((style) => style.styleId);
    for (const style of before.styles) expect(survivors).toContain(style.styleId);
    expect(survivors).toContain('Heading2');
  });

  it('tells a dangling custom reference apart from an undefined built-in', async () => {
    const engine = new Docx4jStylesEngine();
    const dangling = await engine.inspect(await packageReferencing(['MyCustomStyle']), DEFAULT_LIMITS);
    expect(dangling.references.unknownUndefined).toEqual(['MyCustomStyle']);
    expect(dangling.references.builtInUndefined).toEqual([]);

    const builtIn = await engine.inspect(await packageReferencing(['Heading2', 'MyCustomStyle']), DEFAULT_LIMITS);
    // Word synthesizes Heading2 whether or not the document defines it; nothing
    // synthesizes MyCustomStyle, so only that one is a genuinely broken link.
    expect(builtIn.references.builtInUndefined).toEqual(['Heading2']);
    expect(builtIn.references.unknownUndefined).toEqual(['MyCustomStyle']);

    const checked = await engine.verify(await packageReferencing(['MyCustomStyle']), DEFAULT_LIMITS);
    expect(checked.ok).toBe(false);
    expect(checked.checks.find((check) => check.id === 'styles.references.resolve')?.ok).toBe(false);
  });

  it('reports a definition that names properties but carries none', async () => {
    const engine = new Docx4jStylesEngine();
    const blank = await engine.inspect(await packageWithEmptyStyles(), DEFAULT_LIMITS);
    expect(blank.hasStylesPart).toBe(true);
    expect(blank.styles).toEqual([]);
    expect(blank.references.unknownUndefined).toEqual([]);

    const hollow: StyleSpec = { styleId: 'Hollow', name: 'hollow' };
    const checked = await engine.verify(await packageWithEmptyStyles(), DEFAULT_LIMITS, { styles: [hollow] });
    const present = checked.checks.find((check) => check.id === 'styles.definitions.present');
    expect(present?.ok).toBe(false);
  });

  it('writes properties in the sequence ECMA-376 declares', async () => {
    // Out-of-sequence children make Word offer to repair the file: w:pBdr must
    // come before w:spacing, and w:outlineLvl after it.
    const heading = styleElement(HEADINGS[0]!);
    const pBdr = heading.indexOf('<w:pBdr>');
    const spacing = heading.indexOf('<w:spacing');
    const outline = heading.indexOf('<w:outlineLvl');
    expect(pBdr).toBeGreaterThan(-1);
    expect(spacing).toBeGreaterThan(pBdr);
    expect(outline).toBeGreaterThan(spacing);

    const name = heading.indexOf('<w:name');
    const basedOn = heading.indexOf('<w:basedOn');
    const pPr = heading.indexOf('<w:pPr>');
    const rPr = heading.indexOf('<w:rPr>');
    expect(basedOn).toBeGreaterThan(name);
    expect(pPr).toBeGreaterThan(basedOn);
    expect(rPr).toBeGreaterThan(pPr);

    // And the finished part passes the verifier's own order audit.
    const engine = new Docx4jStylesEngine();
    const result = await engine.execute(await richFixture(), HEADINGS, DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS);
    const checked = await engine.verify(result.bytes, DEFAULT_LIMITS, { styles: HEADINGS });
    expect(checked.checks.find((check) => check.id === 'styles.order.valid')?.ok).toBe(true);
  });

  it('writes a border’s w:space in points, the unit ST_PointMeasure declares', () => {
    // Word writes <w:bottom w:val="single" w:sz="4" w:space="4" .../> for a
    // half-point rule offset four points from the text. Carrying `space` through
    // the twips helper asks for a twentieth of that and puts the rule flat
    // against the glyphs instead.
    const heading = styleElement(HEADINGS[0]!);
    expect(heading).toContain('w:space="3"');
    expect(heading).not.toContain('w:space="60"');
  });

  it('rejects malformed plans before touching the document', async () => {
    const engine = new Docx4jStylesEngine();
    const cases: StyleSpec[][] = [
      [],
      [{ styleId: '1bad', name: 'x' }],
      [{ styleId: 'Ok', name: '  ' }],
      [{ styleId: 'Ok', name: 'x' }, { styleId: 'Ok', name: 'y' }],
      [{ styleId: 'Ok', name: 'x', paragraph: { outlineLevel: 9 } }],
      [{ styleId: 'Ok', name: 'x', run: { color: 'FF0000' } }],
      [{ styleId: 'Ok', name: 'x', type: 'paragraph', table: { width: 100 } }],
    ];
    for (const styles of cases) {
      await expect(engine.execute(await richFixture(), styles, DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS))
        .rejects.toMatchObject({ code: expect.stringMatching(/INVALID_INPUT|LIMIT_EXCEEDED/) });
    }
  });

  it('refuses to redefine an existing style when the Profile forbids it', async () => {
    const engine = new Docx4jStylesEngine();
    await expect(engine.execute(await richFixture(), HEADINGS, DEFAULT_LIMITS, {
      ...DEFAULT_FEATURE_FLAGS, allowRedefineExisting: false,
    })).rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION' });

    // Adding a style the document lacks is still allowed under the same flag.
    const added = await engine.execute(await richFixture(), [HEADINGS[1]!], DEFAULT_LIMITS, {
      ...DEFAULT_FEATURE_FLAGS, allowRedefineExisting: false,
    });
    expect(added.defined[0]?.action).toBe('created');
  });

  it('leaves the input bytes untouched', async () => {
    const engine = new Docx4jStylesEngine();
    const source = await richFixture();
    const copy = Uint8Array.from(source);
    await engine.execute(source, HEADINGS, DEFAULT_LIMITS, DEFAULT_FEATURE_FLAGS);
    expect(Array.from(source)).toEqual(Array.from(copy));
  });
});

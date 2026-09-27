import { WordprocessingMLPackage, builtInOf } from '@docx4j/core-ts';
import { StyleDefinitionsPart } from '@docx4j/core-ts/parts';
import type { Element } from '@xmldom/xmldom';
import { DocxStylesError } from '../errors';
import type { StyleFeatureFlags, StyleLimits } from '../contract';
import type {
  DefinedStyle,
  EngineStyleResult,
  ExpectedStyleDefinition,
  StyleInventory,
  StyleSpec,
  StyleSummary,
  StyleType,
} from '../domain/docx-styles';
import { NS, W, attributeValues, childElements, localNameOf, orderProblems, parseXml, serializeXml, xml } from './xml';
import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, STYLE_CHILD_ORDER, TBLPR_CHILD_ORDER, styleElement } from './style-xml';

const DOCX_MAGIC = [0x50, 0x4b, 0x03, 0x04] as const;
const STYLE_TYPES: readonly StyleType[] = ['paragraph', 'character', 'table'];

function assertDocxBytes(bytes: Uint8Array, maxBytes: number): void {
  if (bytes.byteLength > maxBytes) {
    throw new DocxStylesError('LIMIT_EXCEEDED', 'The DOCX exceeds the configured input size limit.');
  }
  if (bytes.byteLength < DOCX_MAGIC.length || !DOCX_MAGIC.every((value, index) => bytes[index] === value)) {
    throw new DocxStylesError('FORMAT_MISMATCH', 'The artifact is not a readable DOCX ZIP package.');
  }
}

async function loadPackage(bytes: Uint8Array, maxBytes: number): Promise<WordprocessingMLPackage> {
  assertDocxBytes(bytes, maxBytes);
  try {
    const loaded = await WordprocessingMLPackage.load(bytes);
    if (!(loaded instanceof WordprocessingMLPackage)) {
      throw new DocxStylesError('FORMAT_MISMATCH', 'The package is not a WordprocessingML document.');
    }
    return loaded;
  } catch (error) {
    if (error instanceof DocxStylesError) throw error;
    throw new DocxStylesError('FORMAT_MISMATCH', 'The DOCX package could not be opened by docx4j-core-ts.');
  }
}

function attr(element: Element, name: string): string | undefined {
  const qualified = element.getAttribute(name);
  if (qualified) return qualified;
  const local = element.getAttributeNS(W, name.replace(/^w:/, ''));
  return local || undefined;
}

function rootElement(text: string): Element {
  const document = parseXml(text);
  const root = document.documentElement;
  if (!root) throw new DocxStylesError('FORMAT_MISMATCH', 'An XML part has no root element.');
  return root as Element;
}

/**
 * Reject a definition that could not be written honestly. Validation runs before
 * the document is touched, so a malformed plan fails on its own terms rather
 * than half-applying and leaving the part inconsistent.
 */
export function assertStyleSpecs(specs: readonly StyleSpec[], limits: StyleLimits, flags: StyleFeatureFlags): void {
  if (specs.length === 0) {
    throw new DocxStylesError('INVALID_INPUT', 'A style plan must define at least one style.');
  }
  if (specs.length > limits.maxStyles) {
    throw new DocxStylesError('LIMIT_EXCEEDED', 'The style plan exceeds the configured style limit.');
  }
  const seen = new Set<string>();
  for (const spec of specs) {
    if (typeof spec.styleId !== 'string' || !/^[A-Za-z_][A-Za-z0-9_.-]{0,254}$/.test(spec.styleId)) {
      throw new DocxStylesError('INVALID_INPUT', `styleId must match the OOXML style-id shape: ${JSON.stringify(spec.styleId)}`);
    }
    if (seen.has(spec.styleId)) {
      throw new DocxStylesError('INVALID_INPUT', `Duplicate styleId in one plan: ${spec.styleId}`);
    }
    seen.add(spec.styleId);
    if (typeof spec.name !== 'string' || spec.name.trim() === '') {
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId} needs a non-empty name.`);
    }
    const type = spec.type ?? 'paragraph';
    if (!STYLE_TYPES.includes(type)) {
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId} has an unsupported type: ${String(type)}`);
    }
    if (!flags.allowStyleDefinitions) {
      throw new DocxStylesError('UNSUPPORTED_OPERATION', 'Style definitions are disabled by Profile configuration.');
    }
    const outline = spec.paragraph?.outlineLevel;
    if (outline !== undefined && (!Number.isInteger(outline) || outline < 0 || outline > 8)) {
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId}: paragraph.outlineLevel is 0-8.`);
    }
    for (const [field, value] of [
      ['paragraph.spacing.before', spec.paragraph?.spacing?.before],
      ['paragraph.spacing.after', spec.paragraph?.spacing?.after],
      ['paragraph.spacing.line', spec.paragraph?.spacing?.line],
      ['run.size', spec.run?.size],
    ] as const) {
      if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
        throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId}: ${field} must be a non-negative number of points.`);
      }
    }
    const color = spec.run?.color;
    if (color !== undefined && !/^(#[0-9a-fA-F]{6}|auto)$/.test(color)) {
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId}: run.color must be '#RRGGBB' or 'auto'.`);
    }
    if (type !== 'table' && spec.table) {
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId}: table properties need type 'table'.`);
    }
    if (type === 'table' && spec.paragraph
      && Object.keys(spec.paragraph).some((key) => key !== 'spacing')) {
      // A table style's `w:pPr` exists to control the rhythm of the paragraphs
      // inside its cells, and `w:spacing` is the property that does it: docx4j's
      // own Table Grid style carries exactly that and nothing else. Wider
      // paragraph properties belong to a paragraph style.
      throw new DocxStylesError('INVALID_INPUT', `Style ${spec.styleId}: a table style's paragraph properties are limited to spacing.`);
    }
  }
}

/** The property groups a spec actually asked for, used for the audit trail. */
function requestedProperties(spec: StyleSpec): string[] {
  const properties: string[] = [];
  const paragraph = spec.paragraph;
  if (paragraph) {
    if (paragraph.keepNext !== undefined) properties.push('paragraph.keepNext');
    if (paragraph.keepLines !== undefined) properties.push('paragraph.keepLines');
    if (paragraph.alignment !== undefined) properties.push('paragraph.alignment');
    if (paragraph.spacing) properties.push('paragraph.spacing');
    if (paragraph.indentation) properties.push('paragraph.indentation');
    if (paragraph.bottomBorder) properties.push('paragraph.bottomBorder');
    if (paragraph.outlineLevel !== undefined) properties.push('paragraph.outlineLevel');
  }
  const run = spec.run;
  if (run) {
    if (run.bold !== undefined) properties.push('run.bold');
    if (run.italic !== undefined) properties.push('run.italic');
    if (run.size !== undefined) properties.push('run.size');
    if (run.color !== undefined) properties.push('run.color');
    if (run.font) properties.push('run.font');
  }
  const table = spec.table;
  if (table) {
    if (table.width !== undefined) properties.push('table.width');
    if (table.layout !== undefined) properties.push('table.layout');
    if (table.cellMargins) properties.push('table.cellMargins');
    if (table.borders) properties.push('table.borders');
  }
  return properties;
}

function summarize(stylesRoot: Element): StyleSummary[] {
  const summaries: StyleSummary[] = [];
  for (const element of childElements(stylesRoot)) {
    if (localNameOf(element) !== 'style') continue;
    const styleId = attr(element, 'w:styleId');
    if (!styleId) continue;
    const nameElement = childElements(element).find((child) => localNameOf(child) === 'name');
    const basedOnElement = childElements(element).find((child) => localNameOf(child) === 'basedOn');
    const carries: string[] = [];
    for (const group of ['pPr', 'rPr', 'tblPr'] as const) {
      const container = childElements(element).find((child) => localNameOf(child) === group);
      if (!container) continue;
      for (const property of childElements(container)) carries.push(`${group}.${localNameOf(property)}`);
    }
    const type = attr(element, 'w:type');
    const summary: StyleSummary = {
      styleId,
      name: nameElement ? (attr(nameElement, 'w:val') ?? styleId) : styleId,
      type: type && STYLE_TYPES.includes(type as StyleType) ? (type as StyleType) : 'paragraph',
      default: ['1', 'true', 'on'].includes((attr(element, 'w:default') ?? '').toLowerCase()),
      carries,
    };
    const basedOn = basedOnElement ? attr(basedOnElement, 'w:val') : undefined;
    if (basedOn) summary.basedOn = basedOn;
    summaries.push(summary);
  }
  return summaries;
}

/** Style ids the content actually references, through all three reference attributes. */
function referencedStyleIds(mainXml: string): string[] {
  const root = rootElement(mainXml);
  return [...new Set([
    ...attributeValues(root, 'pStyle', 'w:val'),
    ...attributeValues(root, 'rStyle', 'w:val'),
    ...attributeValues(root, 'tblStyle', 'w:val'),
  ])];
}

export class Docx4jStylesEngine {
  readonly name = '@docx4j/core-ts';

  /** Read the document's style definitions and what its content references. */
  async inspect(bytes: Uint8Array, limits: StyleLimits): Promise<StyleInventory> {
    const pkg = await loadPackage(bytes, limits.maxInputBytes);
    return this.inventoryOf(pkg);
  }

  private async inventoryOf(pkg: WordprocessingMLPackage): Promise<StyleInventory> {
    const main = pkg.getMainDocumentPart();
    const part = main?.styleDefinitionsPart;
    const styles = part ? summarize(rootElement(await part.getXml())) : [];
    const defined = styles.map((style) => style.styleId);
    const definedSet = new Set(defined);
    const references = main ? referencedStyleIds(await main.getXml()) : [];
    const builtInUndefined: string[] = [];
    const unknownUndefined: string[] = [];
    for (const reference of references) {
      if (definedSet.has(reference)) continue;
      // Word synthesizes the built-in styles it knows even when the part does
      // not define them, so those references resolve to something - just not to
      // anything this document controls.
      if (builtInOf(reference) !== 'Other') builtInUndefined.push(reference);
      else unknownUndefined.push(reference);
    }
    const inventory: StyleInventory = {
      hasStylesPart: part !== undefined,
      styles,
      references: {
        defined: references.filter((reference) => definedSet.has(reference)),
        builtInUndefined,
        unknownUndefined,
      },
    };
    const defaultParagraph = styles.find((style) => style.default && style.type === 'paragraph');
    const defaultTable = styles.find((style) => style.default && style.type === 'table');
    if (defaultParagraph) inventory.defaultParagraphStyle = defaultParagraph.styleId;
    if (defaultTable) inventory.defaultTableStyle = defaultTable.styleId;
    return inventory;
  }

  /**
   * Merge the requested definitions into the styles part and save a new package.
   * Existing definitions the plan does not name are left exactly as they were:
   * a style module that rewrote the whole part would silently drop every style
   * the document already carried.
   */
  async execute(
    bytes: Uint8Array,
    specs: readonly StyleSpec[],
    limits: StyleLimits,
    flags: StyleFeatureFlags,
  ): Promise<EngineStyleResult> {
    assertStyleSpecs(specs, limits, flags);
    const pkg = await loadPackage(bytes, limits.maxInputBytes);
    const main = pkg.getMainDocumentPart();
    if (!main) throw new DocxStylesError('FORMAT_MISMATCH', 'The package has no main document part.');

    const part = main.styleDefinitionsPart ?? new StyleDefinitionsPart();
    if (!main.styleDefinitionsPart) main.addTargetPart(part);
    const existing = main.styleDefinitionsPart ? await part.getXml() : undefined;
    const before = existing ? summarize(rootElement(existing)) : [];
    const beforeIds = new Set(before.map((style) => style.styleId));

    const wanted = new Set(specs.map((spec) => spec.styleId));
    for (const styleId of wanted) {
      if (beforeIds.has(styleId) && !flags.allowRedefineExisting) {
        throw new DocxStylesError(
          'UNSUPPORTED_OPERATION',
          `Style ${styleId} already exists and redefining it requires allowRedefineExisting=true.`,
        );
      }
    }

    const merged = this.merge(existing, specs);
    part.setXml(merged);
    await pkg.refresh();

    const outputBytes = await pkg.save();
    if (outputBytes.byteLength > limits.maxOutputBytes) {
      throw new DocxStylesError('LIMIT_EXCEEDED', 'The edited DOCX exceeds the configured output size limit.');
    }
    const saved = await loadPackage(outputBytes, limits.maxOutputBytes);
    const inventory = await this.inventoryOf(saved);
    const savedIds = new Set(inventory.styles.map((style) => style.styleId));
    for (const spec of specs) {
      if (!savedIds.has(spec.styleId)) {
        throw new DocxStylesError('ENGINE_FAILED', `Style ${spec.styleId} did not survive the save.`, { styleId: spec.styleId });
      }
    }
    const defined: DefinedStyle[] = specs.map((spec) => {
      const action: DefinedStyle['action'] = beforeIds.has(spec.styleId)
        ? (specs.length > 0 ? 'replaced' : 'unchanged')
        : 'created';
      return {
        styleId: spec.styleId,
        name: spec.name,
        type: spec.type ?? 'paragraph',
        action,
        properties: requestedProperties(spec),
      };
    });
    return { bytes: outputBytes, defined, inventory };
  }

  /** Keep every existing definition, replace the named ones, append the new ones. */
  private merge(existingXml: string | undefined, specs: readonly StyleSpec[]): string {
    const wanted = new Set(specs.map((spec) => spec.styleId));
    const appended = specs.map((spec) => styleElement(spec)).join('');
    if (!existingXml) return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>${appended}</w:styles>`;

    const document = parseXml(existingXml);
    const root = document.documentElement as Element | null;
    if (!root || localNameOf(root) !== 'styles') {
      throw new DocxStylesError('FORMAT_MISMATCH', 'The styles part root is not w:styles.');
    }
    for (const child of childElements(root)) {
      if (localNameOf(child) !== 'style') continue;
      const styleId = attr(child, 'w:styleId');
      if (styleId && wanted.has(styleId)) root.removeChild(child);
    }
    const serialized = serializeXml(document);
    const closing = serialized.lastIndexOf('</w:styles>');
    if (closing < 0) throw new DocxStylesError('FORMAT_MISMATCH', 'The styles part has no closing w:styles tag.');
    return serialized.slice(0, closing) + appended + serialized.slice(closing);
  }

  /**
   * Read the saved package back. A definition is only reported present when the
   * part both names it and the content's references to it resolve.
   */
  async verify(
    bytes: Uint8Array,
    limits: StyleLimits,
    expectations?: { styles?: readonly StyleSpec[] },
  ): Promise<{ ok: boolean; checks: { id: string; ok: boolean; message: string }[] }> {
    const pkg = await loadPackage(bytes, limits.maxOutputBytes);
    const inventory = await this.inventoryOf(pkg);
    // The caller states intent with the same spec it wrote; what verification
    // needs is the smaller fact of which property groups were asked for.
    const expected: ExpectedStyleDefinition[] = (expectations?.styles ?? []).map((spec) => ({
      styleId: spec.styleId,
      name: spec.name,
      type: spec.type ?? 'paragraph',
      properties: requestedProperties(spec),
    }));
    const checks: { id: string; ok: boolean; message: string }[] = [];

    checks.push({
      id: 'styles.part.present',
      ok: inventory.hasStylesPart,
      message: inventory.hasStylesPart ? 'The package carries a styles part.' : 'The package has no styles part.',
    });

    const orderIssues: string[] = [];
    if (inventory.hasStylesPart) {
      const main = pkg.getMainDocumentPart();
      const part = main?.styleDefinitionsPart;
      if (part) {
        const root = rootElement(await part.getXml());
        for (const style of childElements(root)) {
          if (localNameOf(style) !== 'style') continue;
          const styleId = attr(style, 'w:styleId') ?? '(unnamed)';
          orderIssues.push(...orderProblems(style, STYLE_CHILD_ORDER, `style ${styleId}`));
          for (const [group, order] of [['pPr', PPR_CHILD_ORDER], ['rPr', RPR_CHILD_ORDER], ['tblPr', TBLPR_CHILD_ORDER]] as const) {
            const container = childElements(style).find((child) => localNameOf(child) === group);
            if (container) orderIssues.push(...orderProblems(container, order, `style ${styleId}/${group}`));
          }
        }
      }
    }
    checks.push({
      id: 'styles.order.valid',
      ok: orderIssues.length === 0,
      message: orderIssues.length === 0
        ? 'Every style writes its properties in the sequence ECMA-376 declares.'
        : `Out-of-sequence properties would make Word offer to repair the file: ${orderIssues.slice(0, 3).join('; ')}`,
    });

    const requested = expected;
    const missing = requested.filter((expected) => !inventory.styles.some((style) => style.styleId === expected.styleId));
    checks.push({
      id: 'styles.definitions.present',
      ok: missing.length === 0,
      message: missing.length === 0
        ? `${requested.length} requested definition(s) are present in the saved part.`
        : `The saved part does not define: ${missing.map((style) => style.styleId).join(', ')}`,
    });

    const blank = requested.filter((expected) => {
      const style = inventory.styles.find((candidate) => candidate.styleId === expected.styleId);
      return style !== undefined && expected.properties.length > 0 && style.carries.length === 0;
    });
    checks.push({
      id: 'styles.definitions.carry',
      ok: blank.length === 0,
      message: blank.length === 0
        ? 'Every requested definition carries the properties it asked for.'
        : `These definitions name properties but carry none: ${blank.map((style) => style.styleId).join(', ')}`,
    });

    checks.push({
      id: 'styles.references.resolve',
      ok: inventory.references.unknownUndefined.length === 0,
      message: inventory.references.unknownUndefined.length === 0
        ? 'Every reference to a non-built-in style resolves to a definition in this package.'
        : `Dangling references to styles nothing defines: ${inventory.references.unknownUndefined.join(', ')}`,
    });
    checks.push({
      id: 'styles.references.builtin',
      ok: true,
      message: inventory.references.builtInUndefined.length === 0
        ? 'The content references no built-in style that this package leaves undefined.'
        : `Renderer-synthesized built-ins are referenced but not defined here: ${inventory.references.builtInUndefined.join(', ')}`,
    });
    return { ok: checks.every((check) => check.ok), checks };
  }
}

/** Re-exported so a caller can name the XML escape the module uses for names. */
export { xml };

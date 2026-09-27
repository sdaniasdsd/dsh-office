import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { computeNodeDigest, createDocxParseModule } from '@dsh-office-profile/docx-parse';
import { DEFAULT_LIMITS } from '../src/config';
import type { DocxEditTarget } from '../src/domain/docx-edit';
import { Docx4jCoreTsEngine } from '../src/engine/adapter';
import { readTableGeometry } from '../src/engine/table-xml';
import { tableTargetFromDualIR, targetFromDualIR } from '../src/mapper';

const fixturePath = fileURLToPath(new URL('../../docx-parse/fixtures/_generated/clean.docx', import.meta.url));

async function fixture(): Promise<Uint8Array> {
  return new Uint8Array(await readFile(fixturePath));
}

function target(): DocxEditTarget {
  const text = 'Hello docx-parse';
  return {
    semanticId: 'test-semantic-id',
    anchor: {
      kind: 'p',
      part: 'word/document.xml',
      structuralPath: '/w:document/w:body/w:p[1]',
      ordinal: 0,
      paraId: null,
      quote: text,
      digest: computeNodeDigest('p', text),
    },
  };
}

describe('docx-edit core-ts adapter', () => {
  it('opens a DOCX and replaces text at a parser-provided paragraph target', async () => {
    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await fixture(), {
      edits: [{ kind: 'replaceText', target: target(), find: 'docx-parse', replace: 'docx-edit' }],
    }, DEFAULT_LIMITS);

    expect(result.edits).toHaveLength(1);
    expect(result.paragraphCount).toBe(1);
    expect(result.bytes.byteLength).toBeGreaterThan(0);
    expect(await engine.verify(result.bytes, DEFAULT_LIMITS, {
      paragraphTexts: ['Hello docx-edit'],
    })).toMatchObject({ ok: true });
  });

  it('uses a real target produced by docx-parse without copying its anchor contract', async () => {
    const parser = createDocxParseModule();
    const parsed = await parser.handlers.execute({
      artifactRef: { id: 'clean-fixture', uri: fixturePath },
      operation: 'execute',
      requestId: 'parse-before-edit',
    });
    const block = parsed.result.ir.content.semantic.blocks.find((item) => item.kind === 'paragraph');
    if (!block || block.kind !== 'paragraph') throw new Error('Expected a paragraph in the parse fixture.');
    const targetRef = targetFromDualIR(parsed.result.ir.content, block.id);
    const result = await new Docx4jCoreTsEngine().execute(await fixture(), {
      edits: [{ kind: 'replaceText', target: targetRef, find: 'docx-parse', replace: 'docx-edit' }],
    }, DEFAULT_LIMITS);

    expect(result.edits[0]?.semanticId).toBe(block.id);
    expect(await new Docx4jCoreTsEngine().verify(result.bytes, DEFAULT_LIMITS, {
      paragraphTexts: ['Hello docx-edit'],
    })).toMatchObject({ ok: true });
    await parser.dispose();
  });

  it('adds a comment to a selected text range and keeps relationships valid', async () => {
    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await fixture(), {
      author: { name: 'docx-edit test', initials: 'DE' },
      edits: [{
        kind: 'addComment',
        target: target(),
        quote: 'docx-parse',
        text: 'Review this term',
        author: { name: 'docx-edit test', initials: 'DE' },
      }],
    }, DEFAULT_LIMITS);

    expect(await engine.verify(result.bytes, DEFAULT_LIMITS, {
      commentTexts: ['Review this term'],
    })).toMatchObject({ ok: true });
  });

  it('fails closed when the parser anchor no longer matches the document', async () => {
    const engine = new Docx4jCoreTsEngine();
    const staleTarget = target();
    staleTarget.anchor = { ...staleTarget.anchor, digest: '0'.repeat(64) };

    await expect(engine.execute(await fixture(), {
      edits: [{ kind: 'replaceText', target: staleTarget, find: 'docx-parse', replace: 'changed' }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  });

  it('writes native tracked changes and can accept them on the same semantic paragraph', async () => {
    const engine = new Docx4jCoreTsEngine();
    const tracked = await engine.execute(await fixture(), {
      edits: [{
        kind: 'replaceText', target: target(), find: 'docx-parse', replace: 'docx-edit', revision: 'track',
      }],
    }, DEFAULT_LIMITS);
    const trackedPackage = await WordprocessingMLPackage.load(tracked.bytes);
    await trackedPackage.getBody();
    expect(trackedPackage.getTrackedChanges()).toHaveLength(2);

    const editedText = 'Hello docx-edit';
    const editedTarget = target();
    editedTarget.anchor = {
      ...editedTarget.anchor,
      quote: editedText,
      digest: computeNodeDigest('p', editedText),
    };
    const accepted = await engine.execute(tracked.bytes, {
      edits: [{ kind: 'resolveRevisions', target: editedTarget, decision: 'accept' }],
    }, DEFAULT_LIMITS);

    expect(accepted.edits[0]?.changed).toBe(2);
    expect(await engine.verify(accepted.bytes, DEFAULT_LIMITS, { paragraphTexts: [editedText] })).toMatchObject({ ok: true });
  });
});

const richFixturePath = fileURLToPath(new URL('../../docx-parse/fixtures/_generated/rich.docx', import.meta.url));

async function richFixture(): Promise<Uint8Array> {
  return new Uint8Array(await readFile(richFixturePath));
}

async function richBlocks() {
  const parser = createDocxParseModule();
  const parsed = await parser.handlers.execute({
    artifactRef: { id: 'rich-fixture', uri: richFixturePath },
    operation: 'execute',
    requestId: 'parse-rich-before-format',
  });
  return { parser, content: parsed.result.ir.content, blocks: parsed.result.ir.content.semantic.blocks };
}

/** ZIP local headers store part names verbatim, so a byte scan is a valid existence check. */
function hasPart(bytes: Uint8Array, name: string): boolean {
  return Buffer.from(bytes).includes(Buffer.from(name, 'utf8'));
}

describe('docx-edit in-place paragraph formatting', () => {
  it('verifies the addressed paragraph when its text is repeated elsewhere in the document', async () => {
    const { parser, content, blocks } = await richBlocks();
    const chapter = blocks.find((block) => block.kind === 'heading' && block.text === 'Chapter One');
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!chapter || !intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [
        { kind: 'replaceText', target: targetFromDualIR(content, chapter.id), find: 'Chapter One', replace: 'Intro text' },
        { kind: 'formatParagraph', target: targetFromDualIR(content, intro.id), font: { bold: true } },
      ],
    }, DEFAULT_LIMITS);
    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, { formats: result.expectedFormats });
    expect(verification.checks.find((check) => check.id === 'format.paragraph.1')).toMatchObject({ ok: true });
    await parser.dispose();
  });

  it('recovers a formatted table-cell paragraph within its cell after an earlier insertion shifts the path', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    const cellParagraph = table?.kind === 'table' ? table.rows[0]?.cells[0]?.paragraphs[0] : undefined;
    if (!table || !cellParagraph) throw new Error('Expected a paragraph in the rich fixture table.');

    const target = targetFromDualIR(content, cellParagraph.id);
    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [
        { kind: 'insertParagraph', target, text: 'Inserted before A1', position: 'Before' },
        { kind: 'formatParagraph', target, font: { bold: true } },
      ],
    }, DEFAULT_LIMITS);
    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, { formats: result.expectedFormats });
    expect(verification.checks.find((check) => check.id === 'format.paragraph.1')).toMatchObject({ ok: true });
    await parser.dispose();
  });

  it('restyles paragraphs and leaves the comments and footnotes parts intact', async () => {
    const { parser, content, blocks } = await richBlocks();
    const chapter = blocks.find((block) => block.kind === 'heading' && block.text === 'Chapter One');
    const section = blocks.find((block) => block.kind === 'heading' && block.text === 'Section');
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!chapter || !section || !intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [
        {
          kind: 'formatParagraph',
          target: targetFromDualIR(content, chapter.id),
          styleId: 'Heading1',
          outlineLevel: 1,
          spaceAfter: 12,
          font: { bold: true, size: 18, color: '#1F3864' },
        },
        {
          kind: 'formatParagraph',
          target: targetFromDualIR(content, section.id),
          styleId: 'Heading2',
          outlineLevel: 2,
          spaceBefore: 12,
          spaceAfter: 6,
          font: { bold: true, size: 14, color: '#2E5496' },
        },
        {
          kind: 'formatParagraph',
          target: targetFromDualIR(content, intro.id),
          spaceAfter: 8,
          lineSpacing: 15,
        },
      ],
    }, DEFAULT_LIMITS);

    // The whole point of formatting in place: the parts a rebuilt package
    // could not carry over are still in the output.
    expect(hasPart(result.bytes, 'word/comments.xml')).toBe(true);
    expect(hasPart(result.bytes, 'word/footnotes.xml')).toBe(true);
    expect(result.expectedFormats).toHaveLength(3);
    expect(result.edits.every((edit) => edit.kind === 'formatParagraph' && edit.changed > 0)).toBe(true);

    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, {
      paragraphTexts: ['Chapter One', 'Section', 'Intro text'],
      commentTexts: ['A review note'],
      formats: result.expectedFormats,
    });
    const formatChecks = verification.checks.filter((check) => check.id.startsWith('format.paragraph.'));
    expect(formatChecks).toHaveLength(3);
    expect(formatChecks.every((check) => check.ok)).toBe(true);
    expect(verification.ok).toBe(true);

    await parser.dispose();
  });

  it('reports the formatting as missing when the saved output lost it', async () => {
    const { parser, content, blocks } = await richBlocks();
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [{ kind: 'formatParagraph', target: targetFromDualIR(content, intro.id), spaceAfter: 8 }],
    }, DEFAULT_LIMITS);

    // Claim a formatting that was never written: verification must catch it.
    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, {
      formats: [{ text: 'Intro text', semanticId: intro.id, bold: true }],
    });
    expect(verification.ok).toBe(false);
    expect(verification.checks.find((check) => check.id === 'format.paragraph.1')?.ok).toBe(false);

    await parser.dispose();
  });

  it('rejects a format request that would change nothing', async () => {
    const engine = new Docx4jCoreTsEngine();
    await expect(engine.execute(await fixture(), {
      edits: [{ kind: 'formatParagraph', target: target() }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
  });

  it('rejects malformed formatting values instead of guessing', async () => {
    const engine = new Docx4jCoreTsEngine();
    const cases = [
      { styleId: '' },
      { outlineLevel: 0 },
      { outlineLevel: 11 },
      { spaceAfter: -1 },
      { font: { color: 'FF0000' } },
      { font: { size: 0 } },
    ] as const;
    for (const request of cases) {
      await expect(engine.execute(await fixture(), {
        edits: [{ kind: 'formatParagraph', target: target(), ...request }],
      }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    }
  });

  it('fails loud on an unrecognised edit kind instead of acting as a sibling kind', async () => {
    const engine = new Docx4jCoreTsEngine();
    // The dispatcher used to end in a bare `else`, so any future or mistyped
    // kind was silently treated as resolveRevisions.
    const plan = {
      edits: [{ kind: 'setPageMargins', target: target() }],
    } as unknown as Parameters<typeof engine.execute>[1];
    await expect(engine.execute(await fixture(), plan, DEFAULT_LIMITS))
      .rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION' });
  });
});

async function mainPartOf(bytes: Uint8Array): Promise<string> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const main = pkg.getMainDocumentPart();
  if (!main) throw new Error('Expected a main document part.');
  return main.getXml();
}

describe('docx-edit table instance geometry', () => {
  it('writes a valid table alignment and verifies it after saving', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    if (!table) throw new Error('Expected the rich fixture table.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [{
        kind: 'formatTable',
        target: tableTargetFromDualIR(content, table.id),
        alignment: 'Center',
        width: 451,
        layout: 'fixed',
        columnWidths: [225.5, 225.5],
      }],
    }, DEFAULT_LIMITS);

    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, {
      tables: result.expectedTableFormats,
    });
    expect(verification.ok).toBe(true);
    expect(verification.checks.find((check) => check.id === 'table.1.geometry')?.ok).toBe(true);
    await parser.dispose();
  });

  it('gives a table with no preferred widths a fixed grid, so its cells stop wrapping', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    if (!table) throw new Error('Expected the rich fixture table.');

    const source = await richFixture();
    // What the fixture carries: a bare grid. A renderer has no preferred width
    // to honour, so it sizes the table to its content and `A1` wraps to two
    // lines. This is the defect the geometry edit exists to fix.
    const before = readTableGeometry(await mainPartOf(source), 0);
    expect(before.present).toBe(true);
    expect(before.gridWidthsTwips).toEqual([0, 0]);
    expect(before.styleId).toBeUndefined();
    expect(before.layout).toBeUndefined();

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(source, {
      edits: [{
        kind: 'formatTable',
        target: tableTargetFromDualIR(content, table.id),
        styleId: 'TableGrid',
        width: 451,
        layout: 'fixed',
        columnWidths: [225.5, 225.5],
        cellMargins: { top: 0, left: 5.4, bottom: 0, right: 5.4 },
        borders: {
          top: { size: 0.5, color: '#9AA6B2' },
          bottom: { size: 0.5, color: '#9AA6B2' },
          insideH: { size: 0.5, color: '#9AA6B2' },
          insideV: { size: 0.5, color: '#9AA6B2' },
        },
        headerRow: true,
        cellVerticalAlignment: 'center',
        rowPagination: [{ rowIndex: 1, cantSplit: true }],
      }],
    }, DEFAULT_LIMITS);

    const after = readTableGeometry(await mainPartOf(result.bytes), 0);
    expect(after.gridWidthsTwips).toEqual([4510, 4510]);
    // Cell widths come back row by row: two rows of two columns.
    expect(after.cellWidthsTwips).toEqual([4510, 4510, 4510, 4510]);
    expect(after.rowCount).toBe(2);
    expect(after.widthTwips).toBe(9020);
    expect(after.layout).toBe('fixed');
    expect(after.styleId).toBe('TableGrid');
    expect(after.hasCellMargins).toBe(true);
    expect(after.hasBorders).toBe(true);
    expect(after.headerRow).toBe(true);
    expect(after.rowCantSplit).toEqual([false, true]);

    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, { tables: result.expectedTableFormats });
    expect(verification.checks.find((check) => check.id === 'table.1.grid')?.ok).toBe(true);
    expect(verification.checks.find((check) => check.id === 'table.1.geometry')?.ok).toBe(true);
    expect(verification.ok).toBe(true);

    await parser.dispose();
  });

  it('reports the grid as unusable when a table still declares no preferred widths', async () => {
    const engine = new Docx4jCoreTsEngine();
    // No geometry was ever written, so a claim that the columns are sized must
    // fail - the point of reading the saved part back instead of trusting the
    // request.
    const verification = await engine.verify(await richFixture(), DEFAULT_LIMITS, {
      tables: [{ tableIndex: 0, columnCount: 2, applied: [], widthPt: 451 }],
    });
    expect(verification.ok).toBe(false);
    expect(verification.checks.find((check) => check.id === 'table.1.grid')?.ok).toBe(false);
    expect(verification.checks.find((check) => check.id === 'table.1.geometry')?.ok).toBe(false);
  });

  it('rejects row pagination settings that do not identify one existing table row', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    if (!table) throw new Error('Expected the rich fixture table.');
    const engine = new Docx4jCoreTsEngine();
    await expect(engine.execute(await richFixture(), {
      edits: [{ kind: 'formatTable', target: tableTargetFromDualIR(content, table.id), rowPagination: [{ rowIndex: 2, cantSplit: true }] }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    await parser.dispose();
  });

  it('refuses a grid that disagrees with the table width, before touching the document', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    if (!table) throw new Error('Expected the rich fixture table.');
    const engine = new Docx4jCoreTsEngine();
    await expect(engine.execute(await richFixture(), {
      edits: [{
        kind: 'formatTable',
        target: tableTargetFromDualIR(content, table.id),
        width: 451,
        columnWidths: [100, 100],
      }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });
    await parser.dispose();
  });

  it('refuses a malformed geometry request and a stray column count', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    if (!table) throw new Error('Expected the rich fixture table.');
    const engine = new Docx4jCoreTsEngine();
    const targetRef = tableTargetFromDualIR(content, table.id);

    await expect(engine.execute(await richFixture(), {
      edits: [{ kind: 'formatTable', target: targetRef }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });

    await expect(engine.execute(await richFixture(), {
      edits: [{ kind: 'formatTable', target: targetRef, width: -1 }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });

    // Two columns in the document, three widths asked for.
    await expect(engine.execute(await richFixture(), {
      edits: [{ kind: 'formatTable', target: targetRef, columnWidths: [100, 100, 100] }],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'INVALID_INPUT' });

    await parser.dispose();
  });

  it('leaves the paragraph edits intact when both kinds are in one plan', async () => {
    const { parser, content, blocks } = await richBlocks();
    const table = blocks.find((block) => block.kind === 'table');
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!table || !intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [
        { kind: 'formatParagraph', target: targetFromDualIR(content, intro.id), spaceAfter: 8, lineSpacing: 15 },
        { kind: 'formatTable', target: tableTargetFromDualIR(content, table.id), width: 451, layout: 'fixed', columnWidths: [225.5, 225.5] },
      ],
    }, DEFAULT_LIMITS);

    // The table pass rewrites the main part, so it has to run after the tree
    // edits rather than from a stale copy of the part.
    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, {
      paragraphTexts: ['Intro text'],
      formats: result.expectedFormats,
      tables: result.expectedTableFormats,
    });
    expect(verification.ok).toBe(true);

    await parser.dispose();
  });
});

describe('docx-edit paragraph insertion', () => {
  it('inserts a paragraph as a Word revision when tracking is requested', async () => {
    const { parser, content, blocks } = await richBlocks();
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [{
        kind: 'insertParagraph',
        target: targetFromDualIR(content, intro.id),
        text: 'Hiring resumed in Q3.',
        position: 'After',
        revision: 'track',
      }],
    }, DEFAULT_LIMITS);

    expect(result.expectedInsertions).toEqual([{ text: 'Hiring resumed in Q3.', tracked: true }]);

    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, { insertions: result.expectedInsertions });
    const check = verification.checks.find((entry) => entry.id === 'insert.1');
    // "The sentence is present" and "the sentence is present as a revision" are
    // different facts; a review workflow needs the second one.
    expect(check?.ok).toBe(true);
    expect(check?.message).toContain('marked as a Word revision');
    expect(verification.ok).toBe(true);

    await parser.dispose();
  });

  it('catches an insertion that is present but carries no revision mark', async () => {
    const { parser, content, blocks } = await richBlocks();
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [{
        kind: 'insertParagraph',
        target: targetFromDualIR(content, intro.id),
        text: 'Plain sentence.',
        position: 'After',
        revision: 'untracked',
      }],
    }, DEFAULT_LIMITS);

    // The text really is there, so claiming it should have been tracked must be
    // reported as a missing revision mark rather than passing on presence.
    const verification = await engine.verify(result.bytes, DEFAULT_LIMITS, {
      insertions: [{ text: 'Plain sentence.', tracked: true }],
    });
    const check = verification.checks.find((entry) => entry.id === 'insert.1');
    expect(check?.ok).toBe(false);
    expect(check?.message).toContain('no Word revision mark');

    await parser.dispose();
  });

  it('rejects an insertion without text or with an unknown side', async () => {
    const { parser, content, blocks } = await richBlocks();
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const targetRef = targetFromDualIR(content, intro.id);
    const cases = [
      { kind: 'insertParagraph', target: targetRef, text: '   ', position: 'After' },
      { kind: 'insertParagraph', target: targetRef, text: 'x', position: 'Sideways' },
    ];
    for (const request of cases) {
      await expect(engine.execute(await richFixture(), { edits: [request as never] }, DEFAULT_LIMITS))
        .rejects.toMatchObject({ code: 'INVALID_INPUT' });
    }

    await parser.dispose();
  });

  it('leaves the target paragraph unchanged and keeps the text in order', async () => {
    const { parser, content, blocks } = await richBlocks();
    const intro = blocks.find((block) => block.kind === 'paragraph' && block.text === 'Intro text');
    if (!intro) throw new Error('Expected the rich fixture blocks.');

    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await richFixture(), {
      edits: [{
        kind: 'insertParagraph',
        target: targetFromDualIR(content, intro.id),
        text: 'Hiring resumed in Q3.',
        position: 'After',
      }],
    }, DEFAULT_LIMITS);

    const mainXml = await mainPartOf(result.bytes);
    const order = ['Intro text', 'Hiring resumed in Q3.', 'Commented paragraph']
      .map((text) => mainXml.indexOf(text));
    expect(order.every((position) => position >= 0)).toBe(true);
    expect(order[1]).toBeGreaterThan(order[0]!);
    expect(order[2]).toBeGreaterThan(order[1]!);

    await parser.dispose();
  });
});

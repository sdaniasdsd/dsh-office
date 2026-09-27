import { describe, expect, it } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { computeNodeDigest } from '@dsh-office-profile/docx-parse';
import { DEFAULT_LIMITS } from '../src/config';
import type { DocxEditTarget } from '../src/domain/docx-edit';
import { Docx4jCoreTsEngine } from '../src/engine/adapter';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * A body paragraph followed by a one-cell table.
 *
 * This is the smallest shape that tells "the container the anchor names" apart
 * from "anywhere in the document": the two paragraphs live in different
 * containers, so a locator that ignores the container cannot distinguish them.
 */
async function outsideAndInside(): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
  const body = [
    '<w:p><w:r><w:t>Outside the table</w:t></w:r></w:p>',
    '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>',
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Inside the cell</w:t></w:r></w:p></w:tc></w:tr>',
    '</w:tbl>',
  ].join('');
  pkg.getMainDocumentPart().setXml(`<w:document ${NS}><w:body>${body}</w:body></w:document>`);
  return pkg.save();
}

function anchorFor(text: string, structuralPath: string): DocxEditTarget {
  return {
    semanticId: `semantic:${text}`,
    anchor: {
      kind: 'p', part: 'word/document.xml', structuralPath, ordinal: 0, paraId: null,
      quote: text, digest: computeNodeDigest('p', text),
    },
  };
}

const edit = (target: DocxEditTarget, find: string) => ({ kind: 'replaceText' as const, target, find, replace: 'EDITED' });

describe('anchor resolution stays inside the container the address names', () => {
  it('refuses an address it cannot walk instead of searching the document for the text', async () => {
    // A path this module does not model (a paragraph inside a content control).
    // Searching the body for text that happens to match would edit a node the
    // anchor never addressed, so the only safe answer is a refusal.
    const engine = new Docx4jCoreTsEngine();
    await expect(engine.execute(await outsideAndInside(), {
      edits: [edit(anchorFor('Outside the table', '/w:document/w:body/w:sdt[1]/w:p[1]'), 'Outside')],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'UNSUPPORTED_OPERATION' });
  });

  it('does not recover a cell target from a paragraph outside its table', async () => {
    // The ordinal drifted past the end of the cell, and the quote exists — but
    // only outside the table the anchor names. Recovering it would edit a
    // paragraph the caller never addressed.
    const engine = new Docx4jCoreTsEngine();
    await expect(engine.execute(await outsideAndInside(), {
      edits: [edit(anchorFor('Outside the table', '/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[9]'), 'Outside')],
    }, DEFAULT_LIMITS)).rejects.toMatchObject({ code: 'STALE_TARGET' });
  });

  it('still recovers from a drifted ordinal when the quote really is in that cell', async () => {
    // The positive control: redundancy is preserved, it is only bounded. The
    // ninth paragraph of a one-paragraph cell does not exist, but the quote
    // does, inside the cell the anchor names.
    const engine = new Docx4jCoreTsEngine();
    const result = await engine.execute(await outsideAndInside(), {
      edits: [edit(anchorFor('Inside the cell', '/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[9]'), 'Inside')],
    }, DEFAULT_LIMITS);

    expect(result.edits).toHaveLength(1);
    expect(await engine.verify(result.bytes, DEFAULT_LIMITS, {
      paragraphTexts: ['Outside the table', 'EDITED the cell'],
    })).toMatchObject({ ok: true });
  });
});

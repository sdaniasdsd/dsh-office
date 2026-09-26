import { readFile, stat } from 'node:fs/promises';
import PDFKit from 'pdfkit';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFName, PDFHexString, PDFString, degrees, rgb } from 'pdf-lib';
import type { PDFFont, PDFPage } from 'pdf-lib';
import type { ArtifactRef, EngineContext, ModuleOptions } from '@dsh-office-profile/pdf-contracts';
import { PdfModuleError, sha256, verifyPolicy } from '@dsh-office-profile/pdf-contracts';
import { createPlanSchema, editPlanSchema, plan } from './plans';
import { inspectPdf } from './native';
import { checkQpdf } from './process';
import { readArtifact } from './process';
import { verifyField } from './forms';

async function fontBytes(context: EngineContext, options: ModuleOptions, ref?: ArtifactRef): Promise<Uint8Array | undefined> {
  if (ref) {
    return readArtifact(ref, context, options);
  }
  if (options.backend?.fontPath) {
    if ((await stat(options.backend.fontPath)).size > context.config.limits.maxInputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Font exceeds read budget.');
    const bytes = await readFile(options.backend.fontPath, { signal: context.signal });
    if (bytes.length > context.config.limits.maxInputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Font exceeds read budget.');
    return bytes;
  }
  return undefined;
}
function ensureGlyphs(text: string, bytes?: Uint8Array) {
  if (!bytes) {
    if (/[^\x09\x0a\x0d\x20-\x7e]/.test(text)) throw new PdfModuleError('FONT_REQUIRED', 'Non-ASCII text requires an explicitly supplied embeddable font.');
    return;
  }
  const font = fontkit.create(bytes);
  for (const character of text) {
    if (/\s/.test(character)) continue;
    if (!font.hasGlyphForCodePoint(character.codePointAt(0)!)) throw new PdfModuleError('FONT_REQUIRED', `Selected font has no glyph for U+${character.codePointAt(0)!.toString(16).toUpperCase()}.`);
  }
}
async function save(context: EngineContext, options: ModuleOptions, bytes: Uint8Array, sources: ArtifactRef[], suggestedName: string) {
  if (!options.writer) throw new PdfModuleError('INVALID_INPUT', 'Artifact writer required.');
  if (bytes.length > context.config.limits.maxOutputBytes) throw new PdfModuleError('LIMIT_EXCEEDED', 'Generated PDF exceeds output budget.');
  const checkContext = { ...context, source: { ref: { id: 'generated', uri: 'memory:generated' }, bytes, identity: { id: 'generated', sha256: sha256(bytes) } } };
  const check = await checkQpdf(checkContext, options);
  if (check.status === 'fail') throw new PdfModuleError('VERIFICATION_FAILED', check.message);
  if (context.request.policy && !verifyPolicy(await inspectPdf(checkContext), context.request.policy).ok)
    throw new PdfModuleError('SAFETY_POLICY_DENIED', 'Generated PDF does not satisfy the declared policy.');
  const reopened = await PDFDocument.load(bytes, { updateMetadata: false });
  if (reopened.getPageCount() > context.config.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'Generated PDF exceeds page budget.');
  context.signal.throwIfAborted();
  const artifactRef = await options.writer.write({ sources, bytes, requestId: context.request.requestId, suggestedName, signal: context.signal });
  return { result: { artifactRef, sources: sources.map(ref => ({ id: ref.id, sha256: ref.sha256! })), visualReview: 'pending' },
    artifacts: [artifactRef], warnings: [{ code: 'VISUAL_REVIEW_PENDING', severity: 'info', message: 'Saved PDF reopened and structurally checked; visual review remains pending.' }] };
}
export async function createPdf(context: EngineContext, options: ModuleOptions) {
  const input = plan(createPlanSchema, context.request.payload);
  const font = await fontBytes(context, options, input.fontRef);
  ensureGlyphs([input.title, ...input.blocks.flatMap(b => b.kind === 'table' ? [...b.headers, ...b.rows.flat()] : 'text' in b ? [b.text] : [])].join('\n'), font);
  const doc = new PDFKit({ size: input.pageSize, margin: 48, bufferPages: true, info: { Title: input.title, Producer: 'dsh-office-profile / PDFKit' } });
  const buffers: Buffer[] = [];
  let size = 0, count = 1;
  const collected = new Promise<Buffer>((resolve, reject) => {
    doc.on('data', (chunk: Buffer) => { size += chunk.length; if (size > context.config.limits.maxOutputBytes) doc.destroy(new PdfModuleError('LIMIT_EXCEEDED', 'PDF stream exceeds output budget.')); else buffers.push(chunk); });
    doc.on('end', () => resolve(Buffer.concat(buffers))); doc.on('error', reject);
  });
  void collected.catch(() => {});
  doc.on('pageAdded', () => { if (++count > context.config.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'Creation exceeds page limit.'); });
  const abort = () => doc.destroy(new Error('aborted'));
  context.signal.addEventListener('abort', abort, { once: true });
  try {
    doc.font(font ? Buffer.from(font) : 'Helvetica');
    const width = doc.page.width - 96;
    for (const block of input.blocks) {
      context.signal.throwIfAborted();
      if (block.kind === 'pageBreak') { doc.addPage(); continue; }
      if (block.kind === 'heading') {
        if (doc.y > doc.page.height - 130) doc.addPage();
        doc.fontSize(block.level === 1 ? 24 : block.level === 2 ? 17 : 13).fillColor('#193B57')
          .text(block.text, 48, doc.y, { width, paragraphGap: 12 });
      } else if (block.kind === 'paragraph') {
        doc.fontSize(block.size).fillColor(block.color).text(block.text, 48, doc.y, { width, paragraphGap: 9, lineGap: 3 });
      } else {
        doc.fontSize(10);
        const cellWidth = width / block.headers.length;
        const rowHeight = (cells: string[]) => Math.max(...cells.map(t => doc.heightOfString(t, { width: cellWidth - 14 }))) + 16;
        const draw = (cells: string[], header: boolean) => {
          const height = rowHeight(cells), y = doc.y;
          if (height > doc.page.height - 120) throw new PdfModuleError('INVALID_INPUT', 'A table row is too tall for one page; split its content.');
          cells.forEach((cell, i) => {
            const x = 48 + i * cellWidth;
            doc.rect(x, y, cellWidth, height).fillAndStroke(header ? '#EAF1F6' : '#FFFFFF', '#C8D5DF');
            doc.fillColor('#243447').text(cell, x + 7, y + 8, { width: cellWidth - 14, height: height - 12 });
          });
          doc.x = 48; doc.y = y + height;
        };
        if (doc.y + rowHeight(block.headers) + (block.rows[0] ? rowHeight(block.rows[0]) : 0) > doc.page.height - 48) doc.addPage();
        draw(block.headers, true);
        for (const row of block.rows) {
          if (doc.y + rowHeight(row) > doc.page.height - 48) { doc.addPage(); draw(block.headers, true); }
          if (doc.y + rowHeight(row) > doc.page.height - 48) throw new PdfModuleError('INVALID_INPUT', 'Table header plus row cannot fit on one page.');
          draw(row, false);
        }
        doc.y += 12;
      }
    }
    const range = doc.bufferedPageRange();
    for (let i = 0; i < range.count; i++) {
      doc.switchToPage(i); doc.page.margins.bottom = 0;
      doc.fontSize(9).fillColor('#627888').text(`${i + 1} / ${range.count}`, 48, doc.page.height - 28, { width, align: 'right', lineBreak: false });
    }
    doc.end();
    return await save(context, options, await collected, input.fontRef ? [input.fontRef] : [], 'created.pdf');
  } catch (error) { doc.destroy(); throw error; }
  finally { context.signal.removeEventListener('abort', abort); }
}
export async function editPdf(context: EngineContext, options: ModuleOptions) {
  const input = plan(editPlanSchema, context.request.payload);
  const profile = await inspectPdf(context);
  if (profile.encrypted || profile.features.signatures || profile.features.xfa)
    throw new PdfModuleError('UNSUPPORTED_OPERATION', 'Editing encrypted, signed or XFA PDFs is not supported by this first editing route.');
  const doc = await PDFDocument.load(context.source!.bytes, { updateMetadata: false });
  for (const operation of input.operations) {
    if (operation.kind !== 'setTextField' && operation.kind !== 'setCheckbox') continue;
    const existing = operation.kind === 'setTextField' ? doc.getForm().getTextField(operation.name).getText() ?? ''
      : doc.getForm().getCheckBox(operation.name).isChecked();
    if (!verifyField(doc, operation.name, existing, false)) throw new PdfModuleError('UNSUPPORTED_OPERATION', 'Field tree and page widgets are missing or ambiguous; explicit repair workflow required.');
  }
  doc.registerFontkit(fontkit);
  const font = await fontBytes(context, options, input.fontRef);
  let embedded: PDFFont | undefined;
  const getFont = async (text: string) => {
    ensureGlyphs(text, font);
    return embedded ??= await doc.embedFont(font ?? 'Helvetica', { subset: true });
  };
  const sources: ArtifactRef[] = [{ ...context.source!.ref, sha256: context.source!.identity.sha256 }];
  if (input.fontRef) sources.push(input.fontRef);
  const pageAt = (n: number) => {
    const page = doc.getPages()[n - 1];
    if (!page) throw new PdfModuleError('INVALID_INPUT', 'Edit references a nonexistent page.');
    return page;
  };
  // Coordinates for stamp/note are CropBox-relative top-left points. Rotated pages
  // are rejected explicitly in this v1 adapter rather than silently mispositioned.
  const point = (page: PDFPage, x: number, y: number) => {
    const unit = page.node.get(PDFName.of('UserUnit'));
    if (page.getRotation().angle % 360 !== 0 || (unit && unit.toString() !== '1'))
      throw new PdfModuleError('UNSUPPORTED_OPERATION', 'Stamp/note requires an unrotated page with UserUnit 1.');
    const crop = page.getCropBox();
    if (x >= crop.width || y >= crop.height) throw new PdfModuleError('INVALID_INPUT', 'Edit point is outside CropBox.');
    return { x: crop.x + x, y: crop.y + crop.height - y, crop };
  };
  for (const operation of input.operations) {
    context.signal.throwIfAborted();
    if (operation.kind === 'rotate') {
      const page = pageAt(operation.page); page.setRotation(degrees((page.getRotation().angle + operation.degrees) % 360));
    } else if (operation.kind === 'selectPages') {
      if (doc.getForm().getFields().length) throw new PdfModuleError('UNSUPPORTED_OPERATION', 'Page selection on forms requires a field/widget remapping workflow.');
      if (new Set(operation.pages).size !== operation.pages.length) throw new PdfModuleError('INVALID_INPUT', 'Page selection cannot duplicate page identities.');
      const pages = operation.pages.map(pageAt);
      while (doc.getPageCount()) doc.removePage(doc.getPageCount() - 1);
      for (const page of pages) doc.addPage(page);
    } else if (operation.kind === 'appendPages') {
      const bytes = await readArtifact(operation.source, context, options);
      const other = await PDFDocument.load(bytes, { updateMetadata: false });
      if (other.getForm().getFields().length) throw new PdfModuleError('UNSUPPORTED_OPERATION', 'Appending interactive form pages is not supported.');
      if (operation.pages.some(p => p > other.getPageCount())) throw new PdfModuleError('INVALID_INPUT', 'Append page out of range.');
      if (doc.getPageCount() + operation.pages.length > context.config.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'Append exceeds page budget before copying.');
      for (const page of await doc.copyPages(other, operation.pages.map(n => n - 1))) doc.addPage(page);
      sources.push({ ...operation.source, sha256: sha256(bytes) });
    } else if (operation.kind === 'stamp') {
      const page = pageAt(operation.page), p = point(page, operation.x, operation.y), font = await getFont(operation.text);
      if (operation.text.includes('\n') || font.widthOfTextAtSize(operation.text, operation.size) + operation.x > p.crop.width || operation.y + operation.size > p.crop.height)
        throw new PdfModuleError('INVALID_INPUT', 'Stamp must fit on one line inside CropBox.');
      const c = operation.color;
      page.drawText(operation.text, { x: p.x, y: p.y - operation.size, size: operation.size, font,
        color: rgb(parseInt(c.slice(1,3),16)/255, parseInt(c.slice(3,5),16)/255, parseInt(c.slice(5,7),16)/255) });
    } else if (operation.kind === 'addNote') {
      const page = pageAt(operation.page), p = point(page, operation.x, operation.y);
      if (operation.x + 20 > p.crop.width || operation.y + 20 > p.crop.height) throw new PdfModuleError('INVALID_INPUT', 'Note icon falls outside CropBox.');
      const note = doc.context.obj({ Type: 'Annot', Subtype: 'Text', Rect: [p.x, p.y - 20, p.x + 20, p.y],
        Contents: PDFHexString.fromText(operation.text), Name: 'Comment', F: 4, P: page.ref, NM: PDFString.of(`note-${sha256(operation.text).slice(0,16)}`) });
      page.node.addAnnot(doc.context.register(note));
    } else if (operation.kind === 'setTextField') {
      const field = doc.getForm().getTextField(operation.name);
      field.setText(operation.value); field.updateAppearances(await getFont(operation.value));
    } else {
      const field = doc.getForm().getCheckBox(operation.name);
      if (operation.checked) field.check(); else field.uncheck(); field.updateAppearances();
    }
    if (doc.getPageCount() > context.config.limits.maxPages) throw new PdfModuleError('LIMIT_EXCEEDED', 'Edited page count exceeds budget.');
  }
  const bytes = await doc.save({ updateFieldAppearances: false });
  // Reopen canonical fields and widget normal appearances before committing.
  const reopened = await PDFDocument.load(bytes, { updateMetadata: false });
  for (const operation of input.operations) {
    if (operation.kind !== 'setTextField' && operation.kind !== 'setCheckbox') continue;
    // Only the last update to a field is the final expectation.
    const last = [...input.operations].reverse().find(o => 'name' in o && o.name === operation.name);
    if (last !== operation) continue;
    if (!verifyField(reopened, operation.name, operation.kind === 'setTextField' ? operation.value : operation.checked))
      throw new PdfModuleError('VERIFICATION_FAILED', 'Written field values/appearances did not survive reopening.');
  }
  return save(context, options, bytes, sources, 'edited.pdf');
}

import { createHash } from 'node:crypto';
import { unzipSync, zipSync } from 'fflate';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { HeaderPart, FooterPart, ImagePart, NumberingDefinitionsPart, StyleDefinitionsPart } from '@docx4j/core-ts/parts';
import type { CreateEngine, EngineRequest, EngineResult } from '../contract';
import { failure } from '../errors';
import { DEFAULT_MARGIN_MM, PAGE_HEIGHT_MM, PAGE_WIDTH_MM } from '../domain/docx-create';
import { NS, xml, textRun, parseXml, serializeXml, W } from './xml';
import { styleXml, PRESETS, headFillColor, footerColorOf } from './presets';
import { REGISTERS, resolveDesign, eighths, twips } from './design';
import { numberingXml, numPrXml } from './numbering';
import { fillTemplate } from './template';

export class Docx4jCreateEngine implements CreateEngine {
  readonly name = 'docx4j-core-ts';
  async execute(input: EngineRequest): Promise<EngineResult> {
    if (input.plan.kind === 'fillTemplate') {
      if (!input.template) failure('INVALID_INPUT', 'Template bytes are missing.');
      const pkg = await WordprocessingMLPackage.load(input.template, { mcePreprocess: false });
      const { filledKeys, changedParts } = await fillTemplate(pkg, input.plan.values);
      const saved = unzipSync(await pkg.save());
      const original = unzipSync(input.template);
      // core-ts saves reachable OPC parts only. Preserve orphan/vendor parts too, and leave
      // content types and relationships byte-identical for this text-only operation.
      for (const name of changedParts) {
        if (!saved[name] || !original[name]) failure('ENGINE_FAILED', 'Template part disappeared during save.');
        original[name] = saved[name];
      }
      return { bytes: zipSync(original), filledKeys };
    }
    const document = input.plan.document;
    // The scenario decides whether this document is decorated at all; the
    // decision travels back to the caller so it can be reviewed rather than
    // trusted. Typography still comes from the preset.
    const design = resolveDesign({
      preset: document.preset,
      ...(document.scenario ? { scenario: document.scenario } : {}),
      ...(document.register ? { register: document.register } : {}),
    });
    const tokens = REGISTERS[design.register];
    const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4', defaultTheme: '2013' });
    const main = pkg.getMainDocumentPart();
    const styles = main.styleDefinitionsPart ?? new StyleDefinitionsPart();
    if (!main.styleDefinitionsPart) main.addTargetPart(styles);
    styles.setXml(styleXml(document.preset, tokens));
    const ids = new Set<string>();
    const para = (key: string, content: string, props = '') => {
      // ST_LongHexNumber: valid positive 31-bit value. Deterministic across text edits.
      const digest = createHash('sha256').update(key).digest();
      const id = ((digest.readUInt32BE(0) & 0x7fffffff) || 1).toString(16).padStart(8, '0').toUpperCase();
      if (ids.has(id)) failure('INVALID_INPUT', 'Paragraph identity collision; rename the block ID.');
      ids.add(id);
      return `<w:p w14:paraId="${id}"><w:pPr>${props}</w:pPr>${content}</w:p>`;
    };
    let refs = '';
    if (document.header !== undefined) {
      const header = new HeaderPart();
      header.setXml(`<w:hdr ${NS}>${para('header', textRun(document.header, '<w:color w:val="667085"/><w:sz w:val="18"/>'), '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="6" w:color="D0D5DD"/></w:pBdr>')}</w:hdr>`);
      refs += `<w:headerReference w:type="default" r:id="${main.addTargetPart(header).id}"/>`;
    }
    if (document.footer !== undefined || document.pageNumbers) {
      const footer = new FooterPart();
      const field = (instruction: string) => `<w:fldSimple w:instr="${instruction}"><w:r><w:t>1</w:t></w:r></w:fldSimple>`;
      // The convention for a bound document states both where the reader is and
      // how much is left, so the page number is not a lone counter.
      const numbers = document.pageNumbers
        ? document.pageNumberStyle === 'pageOfTotal'
          ? `${textRun('第 ')}${field('PAGE')}${textRun(' 页 共 ')}${field('NUMPAGES')}${textRun(' 页')}`
          : field('PAGE')
        : '';
      const runProps = `<w:color w:val="${footerColorOf(document.preset)}"/><w:sz w:val="18"/>`;
      footer.setXml(`<w:ftr ${NS}>${para('footer', textRun(document.footer ?? '', runProps) + (document.footer && numbers ? textRun(' · ') : '') + numbers, '<w:jc w:val="center"/>')}</w:ftr>`);
      refs += `<w:footerReference w:type="default" r:id="${main.addTargetPart(footer).id}"/>`;
    }
    // The package ships an empty numbering part. This module puts definitions
    // in it only when a list block actually needs them, so a document without
    // lists carries no invented numbering.
    if (document.blocks.some((block) => block.kind === 'list')) {
      const numbering = new NumberingDefinitionsPart();
      numbering.setXml(numberingXml(tokens));
      main.addTargetPart(numbering);
    }
    const blocks: string[] = []; let imageId = 0;
    for (const block of document.blocks) {
      const key = `block:${block.id}`;
      switch (block.kind) {
        case 'paragraph':
          blocks.push(para(key, block.runs.map(r => textRun(r.text, `${r.bold ? '<w:b/>' : ''}${r.italic ? '<w:i/>' : ''}${r.color ? `<w:color w:val="${r.color}"/>` : ''}`)).join(''),
            `<w:pStyle w:val="${block.style ?? 'Normal'}"/>${block.keepWithNext ? '<w:keepNext/>' : ''}${block.pageBreakBefore ? '<w:pageBreakBefore/>' : ''}${block.alignment ? `<w:jc w:val="${block.alignment}"/>` : ''}`));
          break;
        case 'list':
          // One paragraph per item, each naming the numbering definition. The
          // marker is not text: a renderer numbers the list and keeps numbering
          // it when items are added or removed later.
          block.items.forEach((item, position) => {
            blocks.push(para(`${key}:${position}`, textRun(item), `<w:pStyle w:val="Normal"/>${numPrXml(block.ordered === true)}`));
          });
          break;
        case 'pageBreak': blocks.push(para(key, '<w:r><w:br w:type="page"/></w:r>')); break;
        case 'toc':
          blocks.push(para(key, textRun(block.title ?? '目录'), '<w:pStyle w:val="Title"/>'));
          blocks.push(para(`${key}:field`, `<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-${block.levels ?? 3}" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${textRun('目录待排版程序更新')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`));
          break;
        case 'table': {
          const widths = block.columnWidthsMm ?? block.rows[0]!.map(() => 160 / block.rows[0]!.length);
          const gridWidths = widths.map(mm => Math.round(mm * 1440 / 25.4));
          const t = tokens.table;
          // A stated absence is `w:val="none"`; omitting the edge would inherit
          // whatever border the table style carries rather than remove it.
          // `borders: 'none'` states the absence of every edge. A layout table —
          // a signature block, a two-column split — is a grid of positions rather
          // than of data, and a visible grid there is noise.
          const border = (['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const)
            .map(side => (block.borders === 'none' || t.borders[side].style === 'none')
              ? `<w:${side} w:val="none"/>`
              : `<w:${side} w:val="single" w:sz="${eighths(t.borders[side].size)}" w:color="${t.borders[side].color}"/>`)
            .join('');
          const margins = (['top', 'left', 'bottom', 'right'] as const)
            .map(side => `<w:${side} w:w="${twips(t.cellMargins[side])}" w:type="dxa"/>`).join('');
          const fill = headFillColor(document.preset, tokens);
          const headRun = `<w:b/>${t.headTextColor ? `<w:color w:val="${t.headTextColor}"/>` : ''}`;
          // Table text sits one step below body text in the registers that say
          // so; `w:sz` is half-points, so a point delta is twice as many units.
          const sizeDelta = t.fontSizeDelta * 2;
          const bodyRun = sizeDelta === 0 ? '' : `<w:sz w:val="${PRESETS[document.preset].bodySize + sizeDelta}"/>`;
          // A document-wide first-line indent is inherited by cell paragraphs,
          // where it is not what "two characters of body indent" means: it eats
          // a cell's width and wraps the label. Reset it — but only when the
          // preset declares an indent, so presets without one keep their exact
          // previous output.
          const cellIndent = PRESETS[document.preset].indent > 0 ? '<w:ind w:firstLineChars="0" w:firstLine="0"/>' : '';
          // A cell carries its own line spacing: a form table is read row by row
          // and must not inherit the body's 1.5-line setting, which makes every
          // row a quarter taller than the row needs.
          const cellSpacing = `<w:spacing w:after="0"${t.cellLineSpacing ? ` w:line="${t.cellLineSpacing}" w:lineRule="auto"` : ''}/>`;
          blocks.push(`<w:tbl><w:tblPr><w:tblW w:w="${gridWidths.reduce((a, v) => a + v, 0)}" w:type="dxa"/><w:tblBorders>${border}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar>${margins}</w:tblCellMar>${t.look ? `<w:tblLook w:val="${t.look}"/>` : ''}</w:tblPr><w:tblGrid>${gridWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${block.rows.map((row, ri) => {
            const head = ri === 0 && block.header !== false;
            const headRule = head && t.headRule
              ? `<w:tcBorders><w:bottom w:val="single" w:sz="${eighths(t.headRule.size)}" w:space="0" w:color="${t.headRule.color}"/></w:tcBorders>`
              : '';
            const vAlign = t.cellVerticalAlignment !== 'top' ? `<w:vAlign w:val="${t.cellVerticalAlignment}"/>` : '';
            return `<w:tr><w:trPr>${head ? '<w:tblHeader/>' : ''}</w:trPr>${row.map((cell, ci) => `<w:tc><w:tcPr><w:tcW w:w="${gridWidths[ci]}" w:type="dxa"/>${headRule}${head && fill ? `<w:shd w:fill="${fill}"/>` : ''}${vAlign}</w:tcPr>${para(`${key}:${ri}:${ci}`, textRun(cell, head ? headRun : bodyRun), `${cellSpacing}${cellIndent}`)}</w:tc>`).join('')}</w:tr>`;
          }).join('')}</w:tbl>`);
          break;
        }
        case 'image': {
          const bytes = input.images[block.id];
          if (!bytes) failure('INVALID_INPUT', 'Image bytes are missing.');
          const png = bytes.length > 24 && Buffer.from(bytes.subarray(0, 8)).equals(Buffer.from([137,80,78,71,13,10,26,10]));
          const jpeg = bytes.length > 4 && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
          if (!png && !jpeg) failure('FORMAT_MISMATCH', 'Only PNG and JPEG images are supported.');
          imageId++;
          const image = new ImagePart(`/word/media/image${imageId}.${png ? 'png' : 'jpg'}`, png ? 'image/png' : 'image/jpeg');
          image.setBytes(bytes); const relId = main.addTargetPart(image).id;
          const cx = Math.round(block.widthMm * 36000); const cy = Math.round(block.heightMm * 36000);
          const drawing = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${imageId}" name="Image ${imageId}" descr="${xml(block.altText)}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${imageId}" name="Image ${imageId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
          blocks.push(para(key, drawing, `${block.caption ? '<w:keepNext/>' : ''}<w:jc w:val="center"/>`));
          if (block.caption) blocks.push(para(`${key}:caption`, textRun(block.caption), '<w:pStyle w:val="Caption"/><w:jc w:val="center"/>'));
          break;
        }
      }
    }
    // Page geometry. Millimetres, the unit the rest of the plan uses, converted
    // here; the historical 25 mm all round is what a plan that says nothing gets.
    const pageSize = document.page?.size ?? 'A4';
    const margin = (edge: 'top' | 'right' | 'bottom' | 'left') =>
      Math.round((document.page?.marginsMm?.[edge] ?? DEFAULT_MARGIN_MM) * 1440 / 25.4);
    main.setXml(`<w:document ${NS}><w:body>${blocks.join('')}<w:sectPr>${refs}<w:pgSz w:w="${Math.round(PAGE_WIDTH_MM[pageSize] * 1440 / 25.4)}" w:h="${Math.round(PAGE_HEIGHT_MM[pageSize] * 1440 / 25.4)}"/><w:pgMar w:top="${margin('top')}" w:right="${margin('right')}" w:bottom="${margin('bottom')}" w:left="${margin('left')}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
    // Request a field refresh, without claiming it has happened.
    if (main.documentSettingsPart) {
      const settings = parseXml(await main.documentSettingsPart.getXml());
      const field = settings.createElementNS(W, 'w:updateFields');
      field.setAttributeNS(W, 'w:val', 'true');
      const existing = settings.getElementsByTagNameNS(W, 'updateFields')[0];
      if (existing) existing.parentNode!.replaceChild(field, existing);
      else {
        // updateFields precedes compatibility, docVars and rsids in CT_Settings.
        const before = Array.from(settings.documentElement!.childNodes).find(n => ['compat', 'docVars', 'rsids', 'mathPr', 'themeFontLang', 'clrSchemeMapping', 'decimalSymbol', 'listSeparator'].includes((n as typeof field).localName ?? ''));
        settings.documentElement!.insertBefore(field, before ?? null);
      }
      main.documentSettingsPart.setXml(serializeXml(settings));
    }
    return { bytes: await pkg.save(), filledKeys: [], design };
  }
  async dispose(): Promise<void> { /* no process or global engine state */ }
}

import type { DocxDualIR, SemanticBlock, SemanticId } from '@dsh-office-profile/docx-parse';
import type { DocxEditTarget } from './domain/docx-edit';
import { DocxEditError } from './errors';

function findBlock(ir: DocxDualIR, semanticId: SemanticId): SemanticBlock | undefined {
  for (const block of ir.semantic.blocks) {
    if (block.id === semanticId) return block;
    if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          const paragraph = cell.paragraphs.find((item) => item.id === semanticId);
          if (paragraph) return paragraph;
        }
      }
    }
  }
  return undefined;
}

/** Resolve a semantic target from the parser's public dual-IR and source-map output. */
export function targetFromDualIR(ir: DocxDualIR, semanticId: SemanticId): DocxEditTarget {
  const block = findBlock(ir, semanticId);
  if (!block) throw new DocxEditError('TARGET_NOT_FOUND', 'The semantic id is not present in the parse result.');
  if (block.anchor.kind !== 'p') {
    throw new DocxEditError('UNSUPPORTED_OPERATION', 'This edit operation requires a paragraph target.');
  }
  const pointers = ir.sourceMap.bySemanticId[semanticId];
  if (!pointers || pointers.length === 0 || !ir.sourceMap.byPointer[pointers[0] ?? '']) {
    throw new DocxEditError('TARGET_NOT_FOUND', 'The semantic node has no physical source-map entry.');
  }
  return { semanticId, anchor: block.anchor };
}

/**
 * Resolve a table target.
 *
 * Kept apart from {@link targetFromDualIR} on purpose: the paragraph path keeps
 * refusing anything that is not a paragraph, so a paragraph edit can never be
 * quietly aimed at a table. A table's geometry is addressed at body level
 * (`/w:document/w:body/w:tbl[n]`), which is what the parser reports for it.
 */
export function tableTargetFromDualIR(ir: DocxDualIR, semanticId: SemanticId): DocxEditTarget {
  const block = findBlock(ir, semanticId);
  if (!block) throw new DocxEditError('TARGET_NOT_FOUND', 'The semantic id is not present in the parse result.');
  if (block.kind !== 'table' || block.anchor.kind !== 'tbl') {
    throw new DocxEditError('UNSUPPORTED_OPERATION', 'A formatTable edit requires a table target.');
  }
  const pointers = ir.sourceMap.bySemanticId[semanticId];
  if (!pointers || pointers.length === 0 || !ir.sourceMap.byPointer[pointers[0] ?? '']) {
    throw new DocxEditError('TARGET_NOT_FOUND', 'The semantic node has no physical source-map entry.');
  }
  return { semanticId, anchor: block.anchor };
}

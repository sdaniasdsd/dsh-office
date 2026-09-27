import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { computeNodeDigest } from '@dsh-office-profile/docx-parse';
import type {
  DocxEditOperation,
  DocxEditTarget,
  EditPlan,
  EngineEditResult,
  ExpectedInsertion,
  ExpectedParagraphFormat,
  ExpectedTableFormat,
  FormatParagraphEdit,
  FormatTableEdit,
  InsertParagraphEdit,
  PackageInspection,
} from '../domain/docx-edit';
import type { EditLimits } from '../contract';
import { DocxEditError } from '../errors';
import { applyTableGeometry, assertTableGeometry, readTableGeometry } from './table-xml';

const DOCX_MAGIC = [0x50, 0x4b, 0x03, 0x04] as const;

/** Word's own bounds: 22 inches of spacing, and its maximum type size. */
const MAX_SPACING_PT = 1584;
const MAX_FONT_SIZE_PT = 1638;

/**
 * A paragraph-format request must name at least one property that can actually
 * change something. Rejecting an empty request here is what keeps a no-op from
 * being reported as a successful edit: the caller gets INVALID_INPUT instead of
 * an artifact that claims formatting it never applied.
 */
function assertFormatRequest(edit: FormatParagraphEdit): void {
  const hasParagraphProperty = edit.styleId !== undefined || edit.outlineLevel !== undefined
    || edit.alignment !== undefined || edit.spaceBefore !== undefined
    || edit.spaceAfter !== undefined || edit.lineSpacing !== undefined;
  const hasFontProperty = edit.font !== undefined
    && Object.values(edit.font).some((value) => value !== undefined);
  if (!hasParagraphProperty && !hasFontProperty) {
    throw new DocxEditError('INVALID_INPUT', 'A formatParagraph edit must set at least one formatting property.');
  }
  if (edit.styleId !== undefined && edit.styleId.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'styleId must be a non-empty string.');
  }
  if (edit.outlineLevel !== undefined
    && (!Number.isInteger(edit.outlineLevel) || edit.outlineLevel < 1 || edit.outlineLevel > 10)) {
    throw new DocxEditError('INVALID_INPUT', 'outlineLevel is 1-9 for heading levels and 10 for body text (Office JS convention).');
  }
  const spacing: readonly (readonly [string, number | undefined])[] = [
    ['spaceBefore', edit.spaceBefore],
    ['spaceAfter', edit.spaceAfter],
    ['lineSpacing', edit.lineSpacing],
  ];
  for (const [name, value] of spacing) {
    if (value !== undefined && (!Number.isFinite(value) || value < 0 || value > MAX_SPACING_PT)) {
      throw new DocxEditError('INVALID_INPUT', `${name} must be a number of points between 0 and ${MAX_SPACING_PT}.`);
    }
  }
  const size = edit.font?.size;
  if (size !== undefined && (!Number.isFinite(size) || size < 1 || size > MAX_FONT_SIZE_PT)) {
    throw new DocxEditError('INVALID_INPUT', `font.size must be a number of points between 1 and ${MAX_FONT_SIZE_PT}.`);
  }
  const color = edit.font?.color;
  if (color !== undefined && !/^#[0-9a-fA-F]{6}$/.test(color)) {
    throw new DocxEditError('INVALID_INPUT', "font.color must be '#RRGGBB'.");
  }
  const name = edit.font?.name;
  if (name !== undefined && name.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'font.name must be a non-empty string.');
  }
}

function assertDocxBytes(bytes: Uint8Array, maxBytes: number): void {
  if (bytes.byteLength > maxBytes) {
    throw new DocxEditError('LIMIT_EXCEEDED', 'The DOCX exceeds the configured input size limit.');
  }
  if (bytes.byteLength < DOCX_MAGIC.length || !DOCX_MAGIC.every((value, index) => bytes[index] === value)) {
    throw new DocxEditError('FORMAT_MISMATCH', 'The artifact is not a readable DOCX ZIP package.');
  }
}

function countComments(comments: Awaited<ReturnType<Awaited<ReturnType<WordprocessingMLPackage['getBody']>>['getComments']>>): number {
  return comments.reduce((total, comment) => total + 1 + comment.replies.length, 0);
}

function relationshipProblems(pkg: WordprocessingMLPackage): string[] {
  const problems: string[] = [];
  const owners = [pkg, ...Array.from(pkg.parts)];
  for (const owner of owners) {
    const relationships = owner.relationshipsPart;
    if (!relationships) continue;
    for (const relationship of relationships.list) {
      if (relationship.targetMode === 'External') continue;
      if (!relationships.getPart(relationship)) {
        problems.push(`${owner.partName.toString()} -> ${relationship.id}`);
      }
    }
  }
  return problems;
}

async function loadPackage(bytes: Uint8Array, maxBytes: number): Promise<WordprocessingMLPackage> {
  assertDocxBytes(bytes, maxBytes);
  try {
    const loaded = await WordprocessingMLPackage.load(bytes);
    if (!(loaded instanceof WordprocessingMLPackage)) {
      throw new DocxEditError('FORMAT_MISMATCH', 'The package is not a WordprocessingML document.');
    }
    return loaded;
  } catch (error) {
    if (error instanceof DocxEditError) throw error;
    throw new DocxEditError('FORMAT_MISMATCH', 'The DOCX package could not be opened by docx4j-core-ts.');
  }
}

/**
 * Every external relationship target in the package, keyed by owner and id.
 *
 * This is what makes it safe to let a document that already carries an external
 * hyperlink through the pre-flight policy: an in-place edit writes text and
 * properties, and this check proves it did not add, drop or retarget a link.
 * A precise invariant after the write is worth more than a coarse refusal
 * before it - the refusal also blocked documents that only needed a number
 * changed.
 */
function externalRelationships(pkg: WordprocessingMLPackage): Map<string, string> {
  const found = new Map<string, string>();
  for (const owner of [pkg, ...Array.from(pkg.parts)]) {
    const relationships = owner.relationshipsPart;
    if (!relationships) continue;
    for (const relationship of relationships.list) {
      if (relationship.targetMode !== 'External') continue;
      found.set(`${owner.partName.toString()}#${relationship.id}`, String(relationship.target));
    }
  }
  return found;
}

/** Human-readable difference between two external-relationship maps. */
function relationshipDiff(before: Map<string, string>, after: Map<string, string>): string[] {
  const differences: string[] = [];
  for (const [key, target] of before) {
    if (!after.has(key)) differences.push(`dropped ${key}`);
    else if (after.get(key) !== target) differences.push(`retargeted ${key}`);
  }
  for (const key of after.keys()) if (!before.has(key)) differences.push(`added ${key}`);
  return differences;
}

export class Docx4jCoreTsEngine {
  readonly name = '@docx4j/core-ts';

  async inspect(bytes: Uint8Array, limits: EditLimits): Promise<PackageInspection> {
    const pkg = await loadPackage(bytes, limits.maxInputBytes);
    const body = await pkg.getBody();
    const outline = await pkg.outline();
    const comments = await body.getComments();
    const trackedChanges = pkg.getTrackedChanges();
    return {
      paragraphCount: outline.paragraphs.length,
      tableCount: outline.tables.length,
      commentCount: countComments(comments),
      trackedChangeCount: trackedChanges.length,
      byteLength: bytes.byteLength,
    };
  }

  async execute(bytes: Uint8Array, plan: EditPlan, limits: EditLimits, allowTrackedChanges = true): Promise<EngineEditResult> {
    if (plan.edits.length === 0) {
      throw new DocxEditError('INVALID_INPUT', 'An edit plan must contain at least one edit.');
    }
    if (plan.edits.length > limits.maxOperations) {
      throw new DocxEditError('LIMIT_EXCEEDED', 'The edit plan exceeds the configured operation limit.');
    }

    // Validate every request shape before the document is touched, so a
    // malformed plan fails on its own terms instead of surfacing as whatever
    // target resolution happens to trip over first.
    for (const edit of plan.edits) {
      if (edit.kind === 'formatParagraph') assertFormatRequest(edit);
      if (edit.kind === 'formatTable') assertTableGeometry(edit);
      if (edit.kind === 'insertParagraph') assertInsertRequest(edit);
    }

    const pkg = await loadPackage(bytes, limits.maxInputBytes);
    const sourceExternal = externalRelationships(pkg);
    const body = await pkg.getBody();
    // Table geometry is written on the part, not through the paragraph views, so
    // it is resolved and applied in its own pass - after the paragraph edits,
    // because it has to run against the document as those edits left it.
    const paragraphEdits = plan.edits.filter((edit) => edit.kind !== 'formatTable');
    const tableEdits = plan.edits.filter((edit): edit is FormatTableEdit => edit.kind === 'formatTable');
    const resolved = new Map<string, (typeof body.paragraphs)[number]>();
    for (const edit of paragraphEdits) {
      if (!resolved.has(edit.target.semanticId)) {
        resolved.set(edit.target.semanticId, this.resolveTarget(body, edit));
      }
    }

    if (plan.author) pkg.author = plan.author;
    const applied = [] as EngineEditResult['edits'];
    const expectedComments: string[] = [];
    const expectedFormats: ExpectedParagraphFormat[] = [];
    const expectedTableFormats: ExpectedTableFormat[] = [];
    const expectedInsertions: ExpectedInsertion[] = [];

    for (const edit of paragraphEdits) {
      const paragraph = resolved.get(edit.target.semanticId);
      if (!paragraph) throw new DocxEditError('TARGET_NOT_FOUND', 'The target paragraph could not be resolved.');
      if (edit.kind === 'replaceText') {
        const mode = edit.revision ?? 'respectDocument';
        const previousMode = await pkg.getChangeTrackingMode();
        if (!allowTrackedChanges && (mode === 'track' || (mode === 'respectDocument' && previousMode !== 'Off'))) {
          throw new DocxEditError('UNSUPPORTED_OPERATION', 'This edit would create tracked changes, which Profile configuration has disabled.');
        }
        const wantedMode = mode === 'track' ? 'TrackAll' : mode === 'untracked' ? 'Off' : previousMode;
        if (wantedMode !== previousMode) await pkg.setChangeTrackingMode(wantedMode);
        try {
          const ranges = paragraph.search(edit.find, { matchCase: edit.matchCase ?? true });
          const range = chooseOccurrence(ranges, edit.occurrence, 'TEXT_NOT_FOUND', 'TEXT_AMBIGUOUS');
          range.insertText(edit.replace, 'Replace');
          applied.push({ kind: edit.kind, semanticId: edit.target.semanticId, changed: 1 });
        } finally {
          if (wantedMode !== previousMode) await pkg.setChangeTrackingMode(previousMode);
        }
      } else if (edit.kind === 'addComment') {
        pkg.author = edit.author;
        if (edit.quote === undefined) {
          await paragraph.insertComment(edit.text);
        } else {
          const ranges = paragraph.search(edit.quote, { matchCase: true });
          const range = chooseOccurrence(ranges, edit.occurrence, 'TEXT_NOT_FOUND', 'TEXT_AMBIGUOUS');
          await range.insertComment(edit.text);
        }
        expectedComments.push(edit.text);
        applied.push({ kind: edit.kind, semanticId: edit.target.semanticId, changed: 1 });
      } else if (edit.kind === 'formatParagraph') {
        const appliedProperties: string[] = [];
        if (edit.styleId !== undefined) {
          paragraph.styleId = edit.styleId;
          appliedProperties.push('styleId');
        }
        if (edit.outlineLevel !== undefined) {
          paragraph.outlineLevel = edit.outlineLevel;
          appliedProperties.push('outlineLevel');
        }
        if (edit.alignment !== undefined) {
          paragraph.alignment = edit.alignment;
          appliedProperties.push('alignment');
        }
        if (edit.spaceBefore !== undefined) {
          paragraph.spaceBefore = edit.spaceBefore;
          appliedProperties.push('spaceBefore');
        }
        if (edit.spaceAfter !== undefined) {
          paragraph.spaceAfter = edit.spaceAfter;
          appliedProperties.push('spaceAfter');
        }
        if (edit.lineSpacing !== undefined) {
          paragraph.lineSpacing = edit.lineSpacing;
          appliedProperties.push('lineSpacing');
        }
        if (edit.font) {
          // `Paragraph.font` writes direct run properties to every run of the paragraph.
          const font = paragraph.font;
          if (edit.font.bold !== undefined) {
            font.bold = edit.font.bold;
            appliedProperties.push('font.bold');
          }
          if (edit.font.italic !== undefined) {
            font.italic = edit.font.italic;
            appliedProperties.push('font.italic');
          }
          if (edit.font.size !== undefined) {
            font.size = edit.font.size;
            appliedProperties.push('font.size');
          }
          if (edit.font.color !== undefined) {
            font.color = edit.font.color;
            appliedProperties.push('font.color');
          }
          if (edit.font.name !== undefined) {
            font.name = edit.font.name;
            appliedProperties.push('font.name');
          }
        }
        // What the caller asked for, recorded so verify() can read the saved
        // package back. Recording the intent is not evidence that it landed.
        const expected: ExpectedParagraphFormat = {
          text: paragraph.text,
          semanticId: edit.target.semanticId,
          ...(edit.target.anchor.structuralPath ? { structuralPath: edit.target.anchor.structuralPath } : {}),
        };
        if (edit.styleId !== undefined) expected.styleId = edit.styleId;
        if (edit.outlineLevel !== undefined) expected.outlineLevel = edit.outlineLevel;
        if (edit.alignment !== undefined) expected.alignment = edit.alignment;
        if (edit.spaceBefore !== undefined) expected.spaceBefore = edit.spaceBefore;
        if (edit.spaceAfter !== undefined) expected.spaceAfter = edit.spaceAfter;
        if (edit.lineSpacing !== undefined) expected.lineSpacing = edit.lineSpacing;
        if (edit.font?.bold !== undefined) expected.bold = edit.font.bold;
        if (edit.font?.italic !== undefined) expected.italic = edit.font.italic;
        if (edit.font?.size !== undefined) expected.size = edit.font.size;
        if (edit.font?.color !== undefined) expected.color = edit.font.color;
        if (edit.font?.name !== undefined) expected.name = edit.font.name;
        expectedFormats.push(expected);
        applied.push({
          kind: edit.kind,
          semanticId: edit.target.semanticId,
          changed: appliedProperties.length,
          details: { applied: appliedProperties.join(',') },
        });
      } else if (edit.kind === 'insertParagraph') {
        const mode = edit.revision ?? 'respectDocument';
        const previousMode = await pkg.getChangeTrackingMode();
        if (!allowTrackedChanges && (mode === 'track' || (mode === 'respectDocument' && previousMode !== 'Off'))) {
          throw new DocxEditError('UNSUPPORTED_OPERATION', 'This edit would create tracked changes, which Profile configuration has disabled.');
        }
        const wantedMode = mode === 'track' ? 'TrackAll' : mode === 'untracked' ? 'Off' : previousMode;
        if (wantedMode !== previousMode) await pkg.setChangeTrackingMode(wantedMode);
        try {
          // The content API asks the package's change tracker for its revision
          // marks, so a tracked insertion lands as `w:ins` on the paragraph mark
          // and its runs rather than as an untracked paragraph.
          const created = paragraph.insertParagraph(edit.text, edit.position);
          if (edit.styleId !== undefined) created.styleId = edit.styleId;
          expectedInsertions.push({ text: edit.text, tracked: wantedMode !== 'Off' });
          applied.push({
            kind: edit.kind,
            semanticId: edit.target.semanticId,
            changed: 1,
            details: { position: edit.position, revision: wantedMode === 'Off' ? 'untracked' : 'tracked' },
          });
        } finally {
          if (wantedMode !== previousMode) await pkg.setChangeTrackingMode(previousMode);
        }
      } else if (edit.kind === 'resolveRevisions') {
        const revisions = paragraph.getTrackedChanges();
        if (revisions.length === 0) {
          throw new DocxEditError('TARGET_NOT_FOUND', 'The target paragraph has no tracked changes to resolve.', {
            semanticId: edit.target.semanticId,
          });
        }
        for (const revision of [...revisions].reverse()) {
          if (edit.decision === 'accept') revision.accept();
          else revision.reject();
        }
        applied.push({ kind: edit.kind, semanticId: edit.target.semanticId, changed: revisions.length, details: { decision: edit.decision } });
      } else {
        // An unrecognised kind must not fall through to a sibling's behavior.
        throw new DocxEditError('UNSUPPORTED_OPERATION', `Unsupported edit kind: ${String((edit as { kind?: unknown }).kind)}`);
      }
    }

    // Second pass: table geometry, on the part as the paragraph edits left it.
    if (tableEdits.length > 0) {
      const main = pkg.getMainDocumentPart();
      if (!main) throw new DocxEditError('FORMAT_MISMATCH', 'The package has no main document part.');
      let mainXml = await main.getXml();
      for (const edit of tableEdits) {
        const tableIndex = tableIndexOf(edit.target);
        const outcome = applyTableGeometry(mainXml, tableIndex, edit);
        mainXml = outcome.xml;
        applied.push({
          kind: edit.kind,
          semanticId: edit.target.semanticId,
          changed: outcome.applied.length,
          details: { applied: outcome.applied.join(','), tableIndex },
        });
        const expected: ExpectedTableFormat = {
          tableIndex,
          columnCount: outcome.columnCount,
          applied: outcome.applied,
        };
        if (edit.styleId !== undefined) expected.styleId = edit.styleId;
        if (edit.width !== undefined) expected.widthPt = edit.width;
        if (edit.layout !== undefined) expected.layout = edit.layout;
        if (edit.alignment !== undefined) expected.alignment = edit.alignment;
        if (edit.columnWidths !== undefined) expected.columnWidthsPt = edit.columnWidths;
        if (edit.headerRow !== undefined) expected.headerRow = edit.headerRow;
        if (edit.rowPagination !== undefined) expected.rowPagination = edit.rowPagination;
        expectedTableFormats.push(expected);
      }
      main.setXml(mainXml);
      await pkg.refresh();
    }

    const expectedParagraphTexts = [...new Set([...resolved.values()].map((paragraph) => paragraph.text))];
    const outputBytes = await pkg.save();
    if (outputBytes.byteLength > limits.maxOutputBytes) {
      throw new DocxEditError('LIMIT_EXCEEDED', 'The edited DOCX exceeds the configured output size limit.');
    }
    const outputPackage = await loadPackage(outputBytes, limits.maxOutputBytes);
    const relationshipDrift = relationshipDiff(sourceExternal, externalRelationships(outputPackage));
    if (relationshipDrift.length > 0) {
      throw new DocxEditError(
        'RELATIONSHIP_INTEGRITY_FAILED',
        `The edit changed the document's external relationships: ${relationshipDrift.slice(0, 3).join('; ')}`,
        { count: relationshipDrift.length },
      );
    }
    const problems = relationshipProblems(outputPackage);
    if (problems.length > 0) {
      throw new DocxEditError('RELATIONSHIP_INTEGRITY_FAILED', 'An internal package relationship no longer resolves.', {
        count: problems.length,
      });
    }
    const outputOutline = await outputPackage.outline();
    return {
      bytes: outputBytes,
      edits: applied,
      touchedParts: [...new Set(plan.edits.map((edit) => edit.kind === 'addComment'
        ? 'word/document.xml + Word comment parts'
        : 'word/document.xml'))],
      paragraphCount: outputOutline.paragraphs.length,
      tableCount: outputOutline.tables.length,
      expectedParagraphTexts,
      expectedCommentTexts: expectedComments,
      expectedFormats,
      expectedTableFormats,
      expectedInsertions,
    };
  }

  async verify(
    bytes: Uint8Array,
    limits: EditLimits,
    expectations?: { paragraphTexts?: string[]; commentTexts?: string[]; formats?: ExpectedParagraphFormat[]; tables?: ExpectedTableFormat[]; insertions?: ExpectedInsertion[] },
  ): Promise<{ ok: boolean; checks: { id: string; ok: boolean; message: string }[] }> {
    const pkg = await loadPackage(bytes, limits.maxOutputBytes);
    const body = await pkg.getBody();
    const paragraphViews = body.paragraphs;
    const paragraphs = paragraphViews.map((paragraph) => paragraph.text);
    const comments = await body.getComments();
    const commentTexts = comments.flatMap((comment) => [comment.content, ...comment.replies.map((reply) => reply.content)]);
    const problems = relationshipProblems(pkg);
    // Read the part back once for the checks that need the raw markup rather
    // than a paragraph view.
    const mainXml = await pkg.getMainDocumentPart()?.getXml() ?? '';
    const checks = [
      { id: 'package.reopens', ok: true, message: 'The output reopens as a WordprocessingML package.' },
      {
        id: 'relationships.resolve',
        ok: problems.length === 0,
        message: problems.length === 0 ? 'All internal relationships resolve.' : `${problems.length} internal relationship(s) do not resolve.`,
      },
      ...((expectations?.paragraphTexts ?? []).map((text, index) => ({
        id: `edit.paragraph.${index + 1}`,
        ok: paragraphs.includes(text),
        message: paragraphs.includes(text) ? 'The edited paragraph is present.' : 'The expected edited paragraph was not found.',
      }))),
      ...((expectations?.commentTexts ?? []).map((text, index) => ({
        id: `edit.comment.${index + 1}`,
        ok: commentTexts.includes(text),
        message: commentTexts.includes(text) ? 'The comment is present.' : 'The expected comment was not found.',
      }))),
      ...formatChecks(body, expectations?.formats ?? []),
      ...tableChecks(mainXml, expectations?.tables ?? []),
      ...insertionChecks(mainXml, expectations?.insertions ?? []),
    ];
    return { ok: checks.every((check) => check.ok), checks };
  }

  /**
   * Resolve the paragraph an anchor names.
   *
   * Anchors are redundant on purpose (docx-parse's composite-anchor-v1 records a
   * paraId, a quote and a structural path precisely because each one fails
   * differently), so a surviving selector may recover a node the others lost.
   * Redundancy is not a licence to search the document, though: the structural
   * path is the scope the caller named as well as an address, so the quote
   * selector is applied inside that scope and widens outward through the
   * containers the path names — never past the table it names, and never at all
   * when the anchor carries no address.
   */
  private resolveTarget(body: Awaited<ReturnType<WordprocessingMLPackage['getBody']>>, edit: DocxEditOperation) {
    const { anchor } = edit.target;
    if (anchor.kind !== 'p') {
      throw new DocxEditError('UNSUPPORTED_OPERATION', 'Only paragraph-level semantic targets can currently be edited.');
    }
    if (anchor.part !== 'word/document.xml') {
      throw new DocxEditError('UNSUPPORTED_OPERATION', 'Only paragraphs in the main document part can currently be edited.');
    }
    const all = body.paragraphs;
    const quoteMatches = (paragraphs: typeof all) =>
      paragraphs.filter((paragraph) => computeNodeDigest('p', paragraph.text) === anchor.digest);

    // 1. A native paraId is an identity, not a hint: it settles the question on
    //    its own, and a text mismatch under a matching paraId is a stale target
    //    rather than a reason to look elsewhere.
    if (anchor.paraId) {
      const byParaId = all.filter((paragraph) => paragraph.paraId?.toUpperCase() === anchor.paraId?.toUpperCase());
      if (byParaId.length > 1) {
        throw new DocxEditError('TARGET_AMBIGUOUS', 'Several paragraphs carry this paraId; no edit was applied.', {
          semanticId: edit.target.semanticId,
          matches: byParaId.length,
        });
      }
      const byIdentity = byParaId[0];
      if (byIdentity) {
        if (computeNodeDigest('p', byIdentity.text) !== anchor.digest) {
          throw new DocxEditError('STALE_TARGET', 'The target paragraph changed after parsing; no edit was applied.', {
            semanticId: edit.target.semanticId,
          });
        }
        return byIdentity;
      }
    }

    // 2. A structural path is an address first and a scope second.
    if (anchor.structuralPath) {
      const scopes = structuralScopes(body, anchor.structuralPath);
      if (!scopes) {
        throw new DocxEditError(
          'UNSUPPORTED_OPERATION',
          'The target names a structural path this module cannot walk; a document-wide text search would edit a node other than the one addressed, so no edit was applied.',
          { semanticId: edit.target.semanticId, structuralPath: anchor.structuralPath },
        );
      }
      const exact = resolveByStructuralPath(body, anchor.structuralPath);
      if (exact && computeNodeDigest('p', exact.text) === anchor.digest) return exact;

      // Paragraphs or rows were added or removed, so the ordinal drifted away
      // from the node the caller meant. The quote still identifies it, but only
      // within the container the address names.
      for (const scope of scopes) {
        const found = quoteMatches(scope.paragraphs);
        if (found.length === 1) return found[0]!;
        if (found.length > 1) {
          throw new DocxEditError('TARGET_AMBIGUOUS', `Several paragraphs inside ${scope.label} match this parse target; no edit was applied.`, {
            semanticId: edit.target.semanticId,
            matches: found.length,
            scope: scope.label,
          });
        }
      }
      throw new DocxEditError('STALE_TARGET', 'The target paragraph is no longer present inside the container the anchor names; no edit was applied.', {
        semanticId: edit.target.semanticId,
        structuralPath: anchor.structuralPath,
      });
    }

    // 3. No address at all: the documented last resort, and the only path
    //    allowed to look across the whole body.
    const found = quoteMatches(all);
    if (found.length === 0) {
      throw new DocxEditError('TARGET_NOT_FOUND', 'The parse target no longer exists in this document.', {
        semanticId: edit.target.semanticId,
      });
    }
    if (found.length > 1) {
      throw new DocxEditError('TARGET_AMBIGUOUS', 'Several paragraphs match this parse target; no edit was applied.', {
        semanticId: edit.target.semanticId,
        matches: found.length,
      });
    }
    return found[0]!;
  }
}

/**
 * The zero-based table index a `tbl` anchor points at.
 *
 * Table geometry is addressed at body level, which is what the parser reports
 * for a table (`/w:document/w:body/w:tbl[n]`). A nested table needs the
 * structural path spelled out before it can be a target; guessing at one would
 * be the kind of fuzzy matching the parser bridge deliberately refuses.
 */
function tableIndexOf(target: DocxEditTarget): number {
  const match = /^\/w:document\/w:body\/w:tbl\[(\d+)\]$/.exec(target.anchor.structuralPath);
  if (!match) {
    throw new DocxEditError(
      'UNSUPPORTED_OPERATION',
      'A formatTable target must name a body-level table (/w:document/w:body/w:tbl[n]).',
      { structuralPath: target.anchor.structuralPath },
    );
  }
  return Number(match[1]) - 1;
}

function resolveByStructuralPath(
  body: Awaited<ReturnType<WordprocessingMLPackage['getBody']>>,
  path: string,
): (typeof body.paragraphs)[number] | undefined {
  const directParagraph = path.match(/^\/w:document\/w:body\/w:p\[(\d+)\]$/);
  if (directParagraph) {
    const ordinal = Number(directParagraph[1]) - 1;
    return body.paragraphs.filter((paragraph) => paragraph.parentTableCell === undefined)[ordinal];
  }

  const cellParagraph = path.match(/^\/w:document\/w:body\/w:tbl\[(\d+)\]\/w:tr\[(\d+)\]\/w:tc\[(\d+)\]\/w:p\[(\d+)\]$/);
  if (!cellParagraph) return undefined;
  const [, tableNumber, rowNumber, cellNumber, paragraphNumber] = cellParagraph;
  const table = body.tables[Number(tableNumber) - 1];
  const row = table?.rows[Number(rowNumber) - 1];
  const cell = row?.cells[Number(cellNumber) - 1];
  return cell?.paragraphs[Number(paragraphNumber) - 1];
}

interface StructuralScope {
  /** Named the way the anchor names it, so a refusal can quote the scope back. */
  label: string;
  paragraphs: Awaited<ReturnType<WordprocessingMLPackage['getBody']>>['paragraphs'];
}

/**
 * The containers a structural path names, narrowest first.
 *
 * Recovery from a drifted ordinal may widen outward through these and no
 * further. A paragraph addressed inside a table cell has no business being
 * recovered from an unrelated part of the body, and an anchor whose path shape
 * this function does not recognise returns `undefined` so the caller can refuse
 * rather than search. A recognised path whose container no longer exists
 * returns an empty list, which is a stale target rather than an unreadable one.
 */
function structuralScopes(
  body: Awaited<ReturnType<WordprocessingMLPackage['getBody']>>,
  path: string,
): StructuralScope[] | undefined {
  if (/^\/w:document\/w:body\/w:p\[\d+\]$/.test(path)) {
    return [{ label: 'the body', paragraphs: body.paragraphs.filter((paragraph) => paragraph.parentTableCell === undefined) }];
  }

  const cellParagraph = /^\/w:document\/w:body\/w:tbl\[(\d+)\]\/w:tr\[(\d+)\]\/w:tc\[(\d+)\]\/w:p\[\d+\]$/.exec(path);
  if (!cellParagraph) return undefined;
  const [, tableNumber, rowNumber, cellNumber] = cellParagraph;
  const table = body.tables[Number(tableNumber) - 1];
  if (!table) return [];
  const tableScope: StructuralScope = {
    label: `w:tbl[${tableNumber}]`,
    paragraphs: table.rows.flatMap((item) => item.cells.flatMap((cellItem) => cellItem.paragraphs)),
  };
  const row = table.rows[Number(rowNumber) - 1];
  if (!row) return [tableScope];
  const rowScope: StructuralScope = {
    label: `w:tbl[${tableNumber}]/w:tr[${rowNumber}]`,
    paragraphs: row.cells.flatMap((cellItem) => cellItem.paragraphs),
  };
  const cell = row.cells[Number(cellNumber) - 1];
  if (!cell) return [rowScope, tableScope];
  return [
    { label: `w:tbl[${tableNumber}]/w:tr[${rowNumber}]/w:tc[${cellNumber}]`, paragraphs: cell.paragraphs },
    rowScope,
    tableScope,
  ];
}

function chooseOccurrence<T>(
  matches: T[],
  occurrence: number | undefined,
  notFoundCode: 'TEXT_NOT_FOUND',
  ambiguousCode: 'TEXT_AMBIGUOUS',
): T {
  if (matches.length === 0) throw new DocxEditError(notFoundCode, 'The requested text was not found in the target paragraph.');
  if (occurrence === undefined && matches.length > 1) {
    throw new DocxEditError(ambiguousCode, 'The requested text occurs more than once; specify its zero-based occurrence.');
  }
  const index = occurrence ?? 0;
  if (!Number.isInteger(index) || index < 0 || index >= matches.length) {
    throw new DocxEditError('INVALID_INPUT', 'The requested occurrence is outside the match range.');
  }
  const selected = matches[index];
  if (!selected) throw new DocxEditError(notFoundCode, 'The requested occurrence was not found.');
  return selected;
}

type BodyView = Awaited<ReturnType<WordprocessingMLPackage['getBody']>>;type ParagraphView = BodyView['paragraphs'][number];

/** Twips round-tripping through points cannot be compared exactly. */
const POINT_TOLERANCE = 0.5;

function approximately(actual: number, expected: number): boolean {
  return Math.abs(actual - expected) <= POINT_TOLERANCE;
}

/**
 * Read the saved package back and assert the requested paragraph formatting is
 * really there. Without this, a formatting edit that silently did nothing would
 * still be reported as a successful `execute`.
 */
function formatChecks(
  body: BodyView,
  formats: readonly ExpectedParagraphFormat[],
): { id: string; ok: boolean; message: string }[] {
  return formats.map((format, index) => {
    const id = `format.paragraph.${index + 1}`;
    let matches: ParagraphView[];
    if (format.structuralPath) {
      const exact = resolveByStructuralPath(body, format.structuralPath);
      if (exact?.text === format.text) {
        matches = [exact];
      } else {
        // If intervening edits shifted the ordinal, recover only inside the
        // containers named by the original address. Never widen to a global
        // text search: legal documents commonly repeat headings in a TOC.
        const scopes = structuralScopes(body, format.structuralPath);
        matches = [];
        for (const scope of scopes ?? []) {
          const recovered = scope.paragraphs.filter((paragraph) => paragraph.text === format.text);
          if (recovered.length > 0) {
            matches = recovered;
            break;
          }
        }
      }
    } else {
      matches = body.paragraphs.filter((paragraph) => paragraph.text === format.text);
    }
    if (matches.length !== 1) {
      return {
        id,
        ok: false,
        message: matches.length === 0
          ? 'The formatted paragraph was not found in the saved output.'
          : 'The formatted paragraph text is ambiguous in the saved output.',
      };
    }
    const paragraph = matches[0]!;
    const missing: string[] = [];
    if (format.styleId !== undefined && paragraph.styleId.toLowerCase() !== format.styleId.toLowerCase()) missing.push('styleId');
    // Read the direct value: the effective read would let a style chain answer
    // for a property this edit never wrote.
    const direct = paragraph.formatting({ direct: true });
    if (format.outlineLevel !== undefined && direct.outlineLevel !== format.outlineLevel) missing.push('outlineLevel');
    if (format.alignment !== undefined && direct.alignment !== format.alignment) missing.push('alignment');
    if (format.spaceBefore !== undefined && !approximately(direct.spaceBefore, format.spaceBefore)) missing.push('spaceBefore');
    if (format.spaceAfter !== undefined && !approximately(direct.spaceAfter, format.spaceAfter)) missing.push('spaceAfter');
    if (format.lineSpacing !== undefined && !approximately(direct.lineSpacing, format.lineSpacing)) missing.push('lineSpacing');
    const font = paragraph.getFont({ direct: true });
    if (format.bold !== undefined && font.bold !== format.bold) missing.push('font.bold');
    if (format.italic !== undefined && font.italic !== format.italic) missing.push('font.italic');
    if (format.size !== undefined && !approximately(font.size, format.size)) missing.push('font.size');
    if (format.color !== undefined && font.color.toLowerCase() !== format.color.toLowerCase()) missing.push('font.color');
    if (format.name !== undefined && font.name.toLowerCase() !== format.name.toLowerCase()) missing.push('font.name');
    return {
      id,
      ok: missing.length === 0,
      message: missing.length === 0
        ? 'The requested paragraph formatting is present in the saved output.'
        : `The saved output does not carry: ${missing.join(', ')}.`,
    };
  });
}

/**
 * A paragraph insertion must name text and a side. Validated before the document
 * is touched, like every other request shape.
 */
function assertInsertRequest(edit: InsertParagraphEdit): void {
  if (typeof edit.text !== 'string' || edit.text.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'An insertParagraph edit needs non-empty text.');
  }
  if (edit.position !== 'Before' && edit.position !== 'After') {
    throw new DocxEditError('INVALID_INPUT', "position must be 'Before' or 'After'.");
  }
  if (edit.styleId !== undefined && edit.styleId.trim() === '') {
    throw new DocxEditError('INVALID_INPUT', 'styleId must be a non-empty string when given.');
  }
}

/**
 * Read the saved part back and assert the insertion is there - and, when a
 * revision was asked for, that it is marked as one. "The sentence is present"
 * and "the sentence is present as a Word revision" are different facts, and the
 * second is what a review workflow actually needs.
 */
function insertionChecks(
  mainXml: string,
  insertions: readonly ExpectedInsertion[],
): { id: string; ok: boolean; message: string }[] {
  const plain = (xml: string): string => xml.replace(/<[^>]+>/g, '');
  return insertions.map((expected, index) => {
    const id = `insert.${index + 1}`;
    if (!mainXml) return { id, ok: false, message: 'The saved main document part could not be read back.' };
    if (!plain(mainXml).includes(expected.text)) {
      return { id, ok: false, message: 'The inserted paragraph text is not present in the saved output.' };
    }
    if (!expected.tracked) return { id, ok: true, message: 'The insertion is present and was written untracked, as requested.' };
    // Revision marks may split the text across runs, so compare the text with
    // the tags stripped rather than against raw markup.
    const markedAsInsertion = [...mainXml.matchAll(/<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/g)]
      .some((match) => plain(match[1] ?? '').includes(expected.text));
    return markedAsInsertion
      ? { id, ok: true, message: 'The insertion is present and marked as a Word revision.' }
      : { id, ok: false, message: 'The inserted text is present but carries no Word revision mark.' };
  });
}

/** Points to the twentieths the format stores. */
const twipsOf = (points: number): number => Math.round(points * 20);

/**
 * Read the saved tables back and assert the geometry is really there.
 *
 * `table.N.grid` is the check that matters for the defect this exists to fix: a
 * table whose every `w:gridCol` reads 0 declares no preferred widths, so a
 * renderer sizes it to its content and a cell like `A1` wraps onto two lines.
 * A style definition cannot repair that, and neither can a claim that the
 * geometry was written - only reading it back can.
 */
function tableChecks(
  mainXml: string,
  tables: readonly ExpectedTableFormat[],
): { id: string; ok: boolean; message: string }[] {
  const checks: { id: string; ok: boolean; message: string }[] = [];
  for (const [position, expected] of tables.entries()) {
    const id = `table.${position + 1}`;
    if (!mainXml) {
      checks.push({ id: `${id}.readable`, ok: false, message: 'The saved main document part could not be read back.' });
      continue;
    }
    const reading = readTableGeometry(mainXml, expected.tableIndex);
    if (!reading.present) {
      checks.push({ id: `${id}.present`, ok: false, message: `The saved package has no table at index ${expected.tableIndex}.` });
      continue;
    }
    const missing: string[] = [];
    if (expected.styleId !== undefined && reading.styleId !== expected.styleId) missing.push('styleId');
    if (expected.widthPt !== undefined && reading.widthTwips !== twipsOf(expected.widthPt)) missing.push('width');
    if (expected.layout !== undefined && reading.layout !== expected.layout) missing.push('layout');
    if (expected.alignment !== undefined && reading.alignment !== expected.alignment.toLowerCase()) missing.push('alignment');
    if (expected.headerRow !== undefined && reading.headerRow !== expected.headerRow) missing.push('headerRow');
    for (const row of expected.rowPagination ?? []) {
      if (reading.rowCantSplit[row.rowIndex] !== row.cantSplit) missing.push(`rowPagination[${row.rowIndex}]`);
    }
    if (expected.columnWidthsPt !== undefined) {
      const wanted = expected.columnWidthsPt.map(twipsOf);
      const gridMatches = reading.gridWidthsTwips.length === wanted.length
        && wanted.every((value, index) => reading.gridWidthsTwips[index] === value);
      if (!gridMatches) missing.push('gridWidths');
      // Cell widths come back row by row, so compare them in chunks of the
      // column count rather than against a flat list.
      const rowMatches = reading.rowCount > 0
        && reading.cellWidthsTwips.length === reading.rowCount * wanted.length
        && reading.cellWidthsTwips.every((value, index) => value === wanted[index % wanted.length]);
      if (!rowMatches) missing.push('cellWidths');
    }
    checks.push({
      id: `${id}.geometry`,
      ok: missing.length === 0,
      message: missing.length === 0
        ? 'The requested table geometry is present in the saved output.'
        : `The saved table does not carry: ${missing.join(', ')}.`,
    });
    const bareColumns = reading.gridWidthsTwips.filter((width) => width <= 0).length;
    checks.push({
      id: `${id}.grid`,
      ok: bareColumns === 0,
      message: bareColumns === 0
        ? 'Every grid column declares a preferred width, so layout does not fall back to content width.'
        : `${bareColumns} of ${reading.gridWidthsTwips.length} grid column(s) declare no preferred width; the table sizes to its content and wraps.`,
    });
  }
  return checks;
}

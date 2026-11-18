// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the toolbar shows (ED-002 section 4.7): the selection's formatting, direct (from the
// marks and attributes) and effective (through CR-001 Phase B's `PropertyResolver`, so a heading
// reads bold because its style says so). Headless, so E1.c only binds.
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { deepCopy, readRunOptions, type Alignment, type Element, type RunFormatting } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { formattingOf, runPropertiesFor, applyParagraphAttrs } from '../export.mjs';
import { paragraphAttrsOf } from '../import.mjs';
import type { ParagraphAttrs } from '../schema.mjs';
import { markSetsIn, atomMarks, ATOM_TYPES, isFormattableAtom } from './marks.mjs';
import { paragraphsIn, points, JC_TO_ALIGNMENT } from './paragraph.mjs';
import { sessionOf, selectionSpans } from './spec.mjs';

export type Mixed<T> = T | 'Mixed';

export interface FontReading {
  bold: Mixed<boolean>;
  italic: Mixed<boolean>;
  underline: Mixed<RunFormatting['underline']>;
  strikeThrough: Mixed<boolean>;
  doubleStrikeThrough: Mixed<boolean>;
  subscript: Mixed<boolean>;
  superscript: Mixed<boolean>;
  name: Mixed<string>;
  size: Mixed<number>;
  color: Mixed<string>;
  highlightColor: Mixed<string | null>;
  style: Mixed<string>;
  rtl: Mixed<boolean>;
}

export interface ParagraphReading {
  style: Mixed<string | null>;
  alignment: Mixed<Alignment | 'Unknown'>;
  /** Points, null when not set. */
  leftIndent: Mixed<number | null>;
  rightIndent: Mixed<number | null>;
  firstLineIndent: Mixed<number | null>;
  spaceBefore: Mixed<number | null>;
  spaceAfter: Mixed<number | null>;
  lineSpacing: Mixed<number | null>;
  bidi: Mixed<boolean>;
}

export interface SelectionFormatting {
  font: FontReading;
  paragraph: ParagraphReading;
}

function agg<T>(values: T[]): Mixed<T> {
  if (values.length === 0) return 'Mixed';
  return values.every((v) => v === values[0]) ? values[0]! : 'Mixed';
}

function fontReading(runs: { fmt: RunFormatting; rtl: boolean }[]): FontReading {
  return {
    bold: agg(runs.map((r) => r.fmt.bold)),
    italic: agg(runs.map((r) => r.fmt.italic)),
    underline: agg(runs.map((r) => r.fmt.underline)),
    strikeThrough: agg(runs.map((r) => r.fmt.strikeThrough)),
    doubleStrikeThrough: agg(runs.map((r) => r.fmt.doubleStrikeThrough)),
    subscript: agg(runs.map((r) => r.fmt.subscript)),
    superscript: agg(runs.map((r) => r.fmt.superscript)),
    name: agg(runs.map((r) => r.fmt.name)),
    size: agg(runs.map((r) => r.fmt.size)),
    color: agg(runs.map((r) => r.fmt.color)),
    highlightColor: agg(runs.map((r) => r.fmt.highlightColor)),
    style: agg(runs.map((r) => r.fmt.style)),
    rtl: agg(runs.map((r) => r.rtl)),
  };
}

function paragraphReading(paragraphs: ParagraphAttrs[]): ParagraphReading {
  const firstLine = (a: ParagraphAttrs): number | null => (a.ind?.hanging != null ? -a.ind.hanging / 20 : points(a.ind?.firstLine ?? null));
  return {
    style: agg(paragraphs.map((a) => a.pStyle)),
    alignment: agg(paragraphs.map((a) => (a.jc ? JC_TO_ALIGNMENT[a.jc] ?? 'Unknown' : 'Unknown'))),
    leftIndent: agg(paragraphs.map((a) => points(a.ind?.start ?? null))),
    rightIndent: agg(paragraphs.map((a) => points(a.ind?.end ?? null))),
    firstLineIndent: agg(paragraphs.map(firstLine)),
    spaceBefore: agg(paragraphs.map((a) => points(a.spacing?.before ?? null))),
    spaceAfter: agg(paragraphs.map((a) => points(a.spacing?.after ?? null))),
    lineSpacing: agg(paragraphs.map((a) => points(a.spacing?.line ?? null))),
    bidi: agg(paragraphs.map((a) => a.bidi)),
  };
}

/** The direct formatting of the selection: what the marks and attributes say, 'Mixed' where they disagree. */
export function selectionFormatting(state: EditorState): SelectionFormatting {
  return {
    font: fontReading(markSetsIn(state).map((marks) => formattingOf(marks))),
    paragraph: paragraphReading(paragraphsIn(state).map((p) => p.node.attrs as ParagraphAttrs)),
  };
}

/** The `w:pPr` a paragraph node stands for: the source paragraph's with the node's attributes applied. */
export function paragraphPropertiesFor(node: PMNode, state: EditorState): wml.PPr | undefined {
  const session = sessionOf(state);
  const attrs = node.attrs as ParagraphAttrs;
  const source = attrs.sourceId && session ? (session.ctx.ids.get(attrs.sourceId) as Element<wml.P> | undefined) : undefined;
  const base = paragraphAttrsOf(source?.value.pPr);
  return applyParagraphAttrs(source?.value.pPr ? deepCopy(source.value.pPr) : undefined, attrs, base);
}

/**
 * The effective formatting of the selection, through the resolver: every run's properties as
 * export would write them, resolved against the paragraph's, the styles and the defaults.
 * Falls back to the direct reading when the session has no resolver.
 */
export function effectiveFormatting(state: EditorState): SelectionFormatting {
  const session = sessionOf(state);
  const resolver = session?.pkg.propertyResolverOrUndefined;
  if (!session || !resolver) return selectionFormatting(state);
  const runs: { fmt: RunFormatting; rtl: boolean }[] = [];
  const { empty } = state.selection;
  const paragraphs = paragraphsIn(state);
  const pPrOf = new Map<PMNode, wml.PPr | undefined>();
  for (const p of paragraphs) pPrOf.set(p.node, paragraphPropertiesFor(p.node, state));
  const resolveRun = (marks: readonly import('prosemirror-model').Mark[], paragraph: PMNode | undefined): void => {
    const runSource = marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const source = runSource ? (session.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
    const rPr = runPropertiesFor(source?.value.rPr, marks);
    const pPr = paragraph ? pPrOf.get(paragraph) : undefined;
    const effective = resolver.getEffectiveRPr(rPr, pPr);
    runs.push({ fmt: readRunOptions(effective), rtl: effective?.rtl !== undefined && effective.rtl.val !== false });
  };
  if (empty) {
    resolveRun(state.storedMarks ?? state.selection.$from.marks(), paragraphs[0]?.node);
  } else {
    let current: PMNode | undefined;
    for (const { from, to } of selectionSpans(state)) state.doc.nodesBetween(from, to, (node) => {
      if (node.type.name === 'paragraph') { current = node; return true; }
      if (node.isText) { resolveRun(node.marks, current); return false; }
      if (ATOM_TYPES.has(node.type.name)) {
        if (!isFormattableAtom(node)) return false;
        const runSource = node.attrs['runSource'] as string | null;
        const marks = atomMarks(node, state);
        const withSource = runSource ? [...marks, state.schema.marks['runSource']!.create({ sourceId: runSource })] : marks;
        resolveRun(withSource, current);
        return false;
      }
      return true;
    });
  }
  const effectiveParagraphs = paragraphs.map((p) => {
    const pPr = resolver.getEffectivePPr(pPrOf.get(p.node));
    const attrs = paragraphAttrsOf(pPr);
    // Effective values have Word's defaults where nothing states them: left, no indent, no space.
    return {
      ...attrs,
      sourceId: null,
      pStyle: (p.node.attrs as ParagraphAttrs).pStyle,
      jc: attrs.jc ?? 'left',
      ind: { start: attrs.ind?.start ?? 0, end: attrs.ind?.end ?? 0, firstLine: attrs.ind?.hanging != null ? null : attrs.ind?.firstLine ?? 0, hanging: attrs.ind?.hanging ?? null },
      spacing: { before: attrs.spacing?.before ?? 0, after: attrs.spacing?.after ?? 0, line: attrs.spacing?.line ?? 240, lineRule: attrs.spacing?.lineRule ?? 'auto' },
    } as ParagraphAttrs;
  });
  return { font: fontReading(runs), paragraph: paragraphReading(effectiveParagraphs) };
}

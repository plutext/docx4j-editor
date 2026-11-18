// SPDX-License-Identifier: GPL-3.0-or-later
//
// Live list labels (ED-003 section 5.4, step C3): the model's node walk (`liveLabels`) held in
// the state, keyed by paragraph node, so the decorations read a label by the node they draw and
// a typed paragraph is numbered like any other. The walk is one `getNumber` per list paragraph
// and runs only on a transaction that changed a paragraph's numbering input (its style or its
// `numPr`) or added or removed a paragraph in a changed range; a keystroke inside a list item's
// text carries the labels forward node by node instead (the steps' maps, as the stylesheet's
// `introducesStyle` reads them).
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { AttrStep, AddMarkStep, RemoveMarkStep, AddNodeMarkStep, RemoveNodeMarkStep } from 'prosemirror-transform';
import type { Node as PMNode } from 'prosemirror-model';
import { liveLabels, levelIndentOf, numberingParts, type EditorDocument, type LiveLabel, type ParagraphAttrs } from '@docx4j/editor-model';

export interface LabelsState {
  labels: Map<PMNode, LiveLabel>;
  /** How many walks so far (tests). */
  walks: number;
}

export const labelsKey = new PluginKey<LabelsState>('docx4j-labels');

/**
 * The meta of a transaction that follows a part edit the document's look depends on (ED-005
 * section 4.4: an Apply in the parts panel, a peer's edit arriving): the labels are walked again
 * and the decorations computed again, though the document did not change.
 */
export const PARTS_CHANGED = 'docx4j-parts-changed';

/**
 * A read-only story's labels, computed outside the view and never walked again: the preview's, from its clone's
 * numbering (ED-005 section 12.117), which the session's walk cannot read.
 */
export function fixedLabelsPlugin(labels: Map<PMNode, LiveLabel>): Plugin<LabelsState> {
  return new Plugin<LabelsState>({ key: labelsKey, state: { init: () => ({ labels, walks: 0 }), apply: (_tr, value) => value } });
}

/** A label by the node drawn, or undefined; the empty lookup for a story. */
export type LabelLookup = (node: PMNode) => LiveLabel | undefined;

/** How many times the labels were computed afresh: unchanged across a transaction means no untouched paragraph's label moved (the changed range's were carried). */
export function labelWalksOf(state: EditorState): number {
  return labelsKey.getState(state)?.walks ?? 0;
}

export function labelLookup(state: EditorState): LabelLookup {
  const labels = labelsKey.getState(state)?.labels;
  return labels ? (node) => labels.get(node) : () => undefined;
}

const rangesOf = new WeakMap<Transaction, [number, number][]>();

/** Touching or overlapping ranges merged, in order. */
function merged(ranges: [number, number][]): [number, number][] {
  ranges.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const out: [number, number][] = [];
  for (const r of ranges) {
    const last = out[out.length - 1];
    if (last && r[0] <= last[1]) last[1] = Math.max(last[1], r[1]);
    else out.push([r[0], r[1]]);
  }
  return out;
}

/**
 * The ranges of the new document a transaction touched (a step's map, or an attribute step's node), merged, and
 * computed once per transaction. Each step's ranges are carried forward through the steps after it, which is work
 * only for a step that moves positions (a replace); a mark or attribute step maps nothing and is skipped, so a
 * tracked deletion of half a document, one mark step per run, costs its steps and not their square (ED-005 section
 * 12.91: that square was most of such a deletion's time, taken by every plugin's gate).
 */
export function changedRanges(tr: Transaction): [number, number][] {
  const cached = rangesOf.get(tr);
  if (cached) return cached;
  let ranges: [number, number][] = [];
  tr.steps.forEach((step, i) => {
    const map = tr.mapping.maps[i]!;
    if ((map as unknown as { ranges: readonly number[] }).ranges.length > 0) {
      for (const r of ranges) { r[0] = map.map(r[0], -1); r[1] = map.map(r[1], 1); }
      map.forEach((_oldStart, _oldEnd, newStart, newEnd) => ranges.push([newStart, newEnd]));
    }
    if (step instanceof AttrStep) ranges.push([step.pos, step.pos + 1]);
    // A mark added or removed maps no positions, so its range is read from the step (ED-005 section
    // 12.10: the review panel's gate missed an accept, which takes a mark off).
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) ranges.push([step.from, step.to]);
    if (step instanceof AddNodeMarkStep || step instanceof RemoveNodeMarkStep) ranges.push([step.pos, step.pos + 1]);
    if (ranges.length > 64) ranges = merged(ranges);
  });
  ranges = merged(ranges);
  rangesOf.set(tr, ranges);
  return ranges;
}

/** The paragraphs of a range, in order. */
function paragraphsIn(doc: PMNode, from: number, to: number): PMNode[] {
  const out: PMNode[] = [];
  doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node) => {
    if (node.type.name === 'paragraph') { out.push(node); return false; }
    return true;
  });
  return out;
}

/** What the count depends on: the style and the paragraph's own `numPr`. */
function numberingInput(node: PMNode): string {
  const a = node.attrs as ParagraphAttrs;
  return JSON.stringify([a.pStyle, a.numPr]);
}

/**
 * Carry the labels across a transaction that changed no numbering input: for each changed
 * range, the paragraphs before and after are paired in order and the old node's label goes to
 * the new one. Undefined when the ranges do not pair (a paragraph added or removed, a style or
 * a `numPr` changed): the caller walks again.
 */
function carried(tr: Transaction, before: PMNode, after: PMNode, labels: Map<PMNode, LiveLabel>, session: EditorDocument): Map<PMNode, LiveLabel> | undefined {
  const out = new Map(labels);
  const inverse = tr.mapping.invert();
  for (const [from, to] of changedRanges(tr)) {
    const news = paragraphsIn(after, from, to);
    const olds = paragraphsIn(before, inverse.map(from, -1), inverse.map(to, 1));
    if (news.length !== olds.length) return undefined;
    for (let i = 0; i < news.length; i++) {
      const n = news[i]!;
      const o = olds[i]!;
      if (n === o) continue;
      if (numberingInput(n) !== numberingInput(o)) return undefined;
      const label = labels.get(o);
      if (label) {
        out.delete(o);
        // The paragraph's own indent is in the label's geometry; a changed one is read again from the level.
        out.set(n, (n.attrs as ParagraphAttrs).ind === (o.attrs as ParagraphAttrs).ind ? label : { ...label, indent: levelIndentOf(n.attrs as ParagraphAttrs, numberingParts(session)) });
      }
    }
  }
  return out;
}

export function labelsPlugin(session: EditorDocument): Plugin<LabelsState> {
  return new Plugin<LabelsState>({
    key: labelsKey,
    state: {
      init: (_config, state) => ({ labels: liveLabels(state.doc, session), walks: 1 }),
      apply: (tr, value, oldState, state) => {
        if (tr.getMeta(PARTS_CHANGED)) return { labels: liveLabels(state.doc, session), walks: value.walks + 1 };
        if (!tr.docChanged) return value;
        const kept = carried(tr, oldState.doc, state.doc, value.labels, session);
        if (kept) return { labels: kept, walks: value.walks };
        return { labels: liveLabels(state.doc, session), walks: value.walks + 1 };
      },
    },
  });
}

/**
 * Whether a transaction's changes touch a node the predicate names, in the new document or the
 * old one (a deletion), by the changed ranges' contents and their ancestors (ED-003 section 8.2:
 * a plugin that walks the whole document for such nodes runs only when one may have changed).
 */
export function touches(tr: Transaction, predicate: (node: PMNode) => boolean): boolean {
  const check = (doc: PMNode, from: number, to: number): boolean => {
    let found = false;
    const $from = doc.resolve(Math.min(from, doc.content.size));
    for (let d = $from.depth; d >= 1 && !found; d--) if (predicate($from.node(d))) found = true;
    if (!found) doc.nodesBetween(from, Math.min(to, doc.content.size), (node) => { if (found) return false; if (predicate(node)) found = true; return !found; });
    return found;
  };
  for (const [from, to] of changedRanges(tr)) if (check(tr.doc, from, to)) return true;
  const invert = tr.mapping.invert();
  for (const [from, to] of changedRanges(tr)) if (check(tr.before, invert.map(from, -1), invert.map(to, 1))) return true;
  return false;
}

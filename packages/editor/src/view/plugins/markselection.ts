// SPDX-License-Identifier: GPL-3.0-or-later
//
// The paragraph marks a selection takes in (ED-005 section 12.81): a paragraph whose end lies inside the
// selection, so that the selection runs on into the next paragraph, has its mark selected as Word has it,
// and its pilcrow is drawn highlighted when the formatting marks are shown (`editor.css`, `.mark-selected`).
// The decorations are those paragraphs alone, read again when the selection or the document changes:
// a caret or a selection inside one paragraph makes none, and a selection over the whole of a long
// document is capped, since a decoration per block is the one cost the view must not carry.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';

export const MARK_SELECTED_CLASS = 'mark-selected';
/** The most paragraphs decorated at once: beyond it (a select-all over hundreds of pages) the marks go unpainted. */
export const MARK_SELECTION_CAP = 400;

const key = new PluginKey<DecorationSet>('markSelection');

/** The paragraphs whose mark the selection `from`..`to` takes in: their end is at or after `from` and before `to`. */
export function markSelectedRanges(doc: PMNode, from: number, to: number): { from: number; to: number }[] {
  const out: { from: number; to: number }[] = [];
  if (to <= from) return out;
  doc.nodesBetween(from, to, (node, pos) => {
    if (out.length >= MARK_SELECTION_CAP) return false;
    if (!node.isTextblock) return true;
    const end = pos + node.nodeSize - 1;
    if (end >= from && end < to) out.push({ from: pos, to: pos + node.nodeSize });
    return false;
  });
  return out.length >= MARK_SELECTION_CAP ? [] : out;
}

function decorationsOf(doc: PMNode, from: number, to: number): DecorationSet {
  const ranges = markSelectedRanges(doc, from, to);
  return ranges.length === 0 ? DecorationSet.empty : DecorationSet.create(doc, ranges.map((r) => Decoration.node(r.from, r.to, { class: MARK_SELECTED_CLASS })));
}

export function markSelectionPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: (_config, state) => decorationsOf(state.doc, state.selection.from, state.selection.to),
      apply: (tr, set, _old, state) => (tr.selectionSet || tr.docChanged ? decorationsOf(state.doc, state.selection.from, state.selection.to) : set),
    },
    props: { decorations: (state) => key.getState(state) ?? DecorationSet.empty },
  });
}

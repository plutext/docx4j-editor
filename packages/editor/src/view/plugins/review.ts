// SPDX-License-Identifier: GPL-3.0-or-later
//
// The review panel's current change on the page (ED-005 section 6.5): an inline decoration over
// each of its ranges, both halves of a move since ProseMirror has one selection range, set by a
// transaction's meta and mapped through the rest; never one per change, and none when nothing is
// current. A range that no longer holds a revision after a change to the document (the change
// accepted or rejected, here or by a peer) drops its decoration (ED-005 section 12.105): the tint
// had stayed on accepted text until another entry was selected.
import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { Decoration, DecorationSet } from 'prosemirror-view';

const REVISION_MARKS = new Set(['insertion', 'deletion', 'formatChange']);

/** Whether `[from, to]` still holds a revision: a mark on its text, or the revision of a node it covers whole (a row, an atom, a paragraph's mark). */
function holdsRevision(doc: PMNode, from: number, to: number): boolean {
  let found = false;
  doc.nodesBetween(from, to, (node, pos) => {
    if (found) return false;
    if (node.isText) { if (node.marks.some((m) => REVISION_MARKS.has(m.type.name))) found = true; return false; }
    if (to > pos + node.nodeSize - 1) {
      const a = node.attrs as Record<string, unknown>;
      const changed = (key: string, inner: string): boolean => !!(a[key] as Record<string, unknown> | null | undefined)?.[inner];
      if (a['revision'] || a['markRevision'] || a['pPrChange'] || changed('trPr', 'trPrChange') || changed('tcPr', 'tcPrChange') || changed('markRPr', 'rPrChange') || changed('sectPr', 'sectPrChange')) found = true;
    }
    return !found;
  });
  return found;
}

export const currentChangeKey = new PluginKey<DecorationSet>('docx4j-current-change');

/** The meta that sets the current change's ranges, or clears them with null. */
export function setCurrentChange(tr: Transaction, ranges: readonly [number, number][] | null): Transaction {
  return tr.setMeta(currentChangeKey, ranges);
}

export function currentChangePlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: currentChangeKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, set) {
        const ranges = tr.getMeta(currentChangeKey) as readonly [number, number][] | null | undefined;
        if (ranges === undefined) {
          const mapped = set.map(tr.mapping, tr.doc);
          if (!tr.docChanged || mapped === DecorationSet.empty) return mapped;
          const stale = mapped.find().filter((d) => !holdsRevision(tr.doc, d.from, d.to));
          return stale.length === 0 ? mapped : mapped.remove(stale);
        }
        if (ranges === null) return DecorationSet.empty;
        return DecorationSet.create(tr.doc, ranges.filter(([from, to]) => from < to).map(([from, to]) => Decoration.inline(from, to, { class: 'current-change' })));
      },
    },
    props: {
      decorations: (state) => currentChangeKey.getState(state),
    },
  });
}

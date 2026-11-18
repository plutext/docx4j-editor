// SPDX-License-Identifier: GPL-3.0-or-later
//
// UPSTREAM-15 and 16: y-prosemirror's diff takes a text whose attribute-less mark (bold, italic, a run's
// direction) went as another mark came for unchanged, so the store kept the old formatting: with
// tracking on, bold taken off text (which gains a `formatChange` for it) was lost at the save and
// never reached a peer. The model's `syncTextFormatting` writes what the diff misjudged; here each
// surface's binding runs it after its own write, in the same Yjs transaction (one update, one undo
// step), over the top-level blocks the write can have changed: those whose node is not the one the
// binding's mapping holds for them, found from each end as y-prosemirror's own diff finds them. Before
// the write, those blocks' children are made to match the new ones' in kind, so that the diff never
// replaces a child with a copy of its neighbour (UPSTREAM-16).
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type * as Y from 'yjs';
import { ySyncPluginKey, type ProsemirrorBinding } from 'y-prosemirror';
import { syncTextFormatting, alignBlocks } from '@docx4j/editor-model';

/** Have the view's sync binding write the text formatting its diff misjudges (UPSTREAM-15); a view without a binding is left alone. */
export function syncFormattingOnWrite(view: EditorView): void {
  const binding = (ySyncPluginKey.getState(view.state) as { binding?: ProsemirrorBinding } | undefined)?.binding;
  if (!binding) return;
  const write = binding._prosemirrorChanged.bind(binding);
  const fragment = binding.type as Y.XmlFragment;
  const mapping = binding.mapping as Map<unknown, unknown>;
  binding._prosemirrorChanged = (doc: PMNode) => {
    const ys = fragment.toArray();
    const n = doc.childCount;
    let from = 0;
    while (from < n && from < ys.length && mapping.get(ys[from]) === doc.child(from)) from++;
    let to = n;
    let toY = ys.length;
    while (to > from && toY > from && mapping.get(ys[toY - 1]) === doc.child(to - 1)) { to--; toY--; }
    fragment.doc!.transact(() => {
      // UPSTREAM-16: the changed blocks made to match the new ones in kind first, down the tree.
      if (from < toY) alignBlocks(fragment, doc, from, to, toY, (i) => mapping.get(ys[i]), mapping);
      write(doc);
      if (from < to) syncTextFormatting(fragment, doc, { from, to });
    }, ySyncPluginKey);
  };
}

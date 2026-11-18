// SPDX-License-Identifier: GPL-3.0-or-later
//
// Remote updates as small transactions (ED-004 section 6.3). y-prosemirror applies every change
// that arrives from Yjs, a colleague's keystroke included, by replacing the whole document with
// the fragment's content (`ProsemirrorBinding._typeChanged`: `this._tr.replace(0, size, …)`).
// The unchanged blocks are the same node objects (the binding's mapping caches them), so the
// document ends up right; but the transaction's one step spans the whole document, so every
// plugin gated on the range a transaction touched (ED-003 section 8.2: labels, row groups,
// decorations, tab stops) takes a colleague's keystroke for a change everywhere, and a remote
// keystroke cost 35 to 41 ms on the 255-page document where a local one costs 9 (ED-004 section
// 10.7 item 4).
//
// UPSTREAM-9: the binding builds its transactions through a `_tr` getter. Each surface's binding
// is given one whose whole-document `replace` is narrowed to the blocks that differ, by identity,
// between the document and the new content: the common leading and trailing blocks are left out
// of the step. The resulting document is the same node for node; only the step is smaller.
import { Fragment, Slice } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { Transaction } from 'prosemirror-state';
import { ySyncPluginKey } from 'y-prosemirror';

/** A replace of the whole document narrowed to the top-level blocks that are not the same objects. */
export function narrowed(tr: Transaction): Transaction {
  const replace = tr.replace.bind(tr);
  tr.replace = ((from: number, to: number = from, slice: Slice = Slice.empty) => {
    const doc = tr.doc;
    if (from !== 0 || to !== doc.content.size || slice.openStart !== 0 || slice.openEnd !== 0) return replace(from, to, slice);
    const next = slice.content;
    const n = doc.childCount;
    const m = next.childCount;
    let start = 0;
    while (start < n && start < m && doc.child(start) === next.child(start)) start++;
    let endOld = n;
    let endNew = m;
    while (endOld > start && endNew > start && doc.child(endOld - 1) === next.child(endNew - 1)) { endOld--; endNew--; }
    if (start === endOld && start === endNew) return tr;   // the same blocks: nothing to replace
    let fromPos = 0;
    for (let i = 0; i < start; i++) fromPos += doc.child(i).nodeSize;
    let toPos = fromPos;
    for (let i = start; i < endOld; i++) toPos += doc.child(i).nodeSize;
    const middle = [];
    for (let i = start; i < endNew; i++) middle.push(next.child(i));
    return replace(fromPos, toPos, new Slice(Fragment.from(middle), 0, 0));
  }) as Transaction['replace'];
  return tr;
}

/** Give the view's sync binding narrowed transactions (UPSTREAM-9); a view without a binding is left alone. */
export function narrowRemoteReplaces(view: EditorView): void {
  const binding = (ySyncPluginKey.getState(view.state) as { binding?: object } | undefined)?.binding;
  if (!binding) return;
  Object.defineProperty(binding, '_tr', {
    configurable: true,
    get: () => narrowed(view.state.tr.setMeta('addToHistory', false)),
  });
}

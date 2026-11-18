// SPDX-License-Identifier: GPL-3.0-or-later
//
// Zero-width atoms (bookmarks, field chars, hidden deletions, page-break markers, empty runs)
// are caret positions ProseMirror would stop at (ED-002 section 10.1 item 3). ArrowLeft and
// ArrowRight step over a run of them in one press, and a selection edge that lands on one moves
// on to the nearest text position.
//
// A selection that covers a paragraph's visible text takes the invisible atoms at its edges with
// it (ED-010 section 9.3, Jason's cut of a heading whose nine `_Toc` bookmarks start before its
// first character): a selection made by the mouse or the selection bar starts at the first visible
// position, after the markers, so a cut left the markers behind and the paste found their names
// taken. Word takes a bookmark whose whole text is selected along with the text. A selection
// that covers part of a paragraph's text is left as it is: its markers stay, as Word shrinks a
// bookmark cut in part.
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode, ResolvedPos } from 'prosemirror-model';

export const invisibleAtomsKey = new PluginKey('docx4j-invisible-atoms');

/** Kinds of `opaqueInline` that render nothing. */
export const INVISIBLE_KINDS = new Set(['bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd', 'fldChar', 'instrText', 'del', 'moveFrom', 'lastRenderedPageBreak', 'r', 'proofErr', 'commentRangeStart', 'commentRangeEnd']);

export function isInvisible(node: PMNode | null | undefined): boolean {
  return !!node && node.type.name === 'opaqueInline' && INVISIBLE_KINDS.has(String(node.attrs['kind']));
}

/** The position after skipping invisible atoms from `pos` in `dir`; `pos` when there are none. */
export function skipInvisible(state: EditorState, pos: number, dir: -1 | 1): number {
  let $pos = state.doc.resolve(pos);
  for (;;) {
    const next = dir > 0 ? $pos.nodeAfter : $pos.nodeBefore;
    if (!isInvisible(next)) return $pos.pos;
    $pos = state.doc.resolve($pos.pos + dir * next!.nodeSize);
  }
}

/** The first position in a textblock after its leading invisible atoms, and the last before its trailing ones. */
function visibleBounds($pos: ResolvedPos): { start: number; end: number } {
  const parent = $pos.parent;
  const start = $pos.start();
  let first = 0;
  while (first < parent.childCount && isInvisible(parent.child(first))) first++;
  let last = parent.childCount;
  while (last > first && isInvisible(parent.child(last - 1))) last--;
  let a = start;
  for (let i = 0; i < first; i++) a += parent.child(i).nodeSize;
  let b = start + parent.content.size;
  for (let i = parent.childCount - 1; i >= last; i--) b -= parent.child(i).nodeSize;
  return { start: a, end: b };
}

/**
 * A text selection widened over the invisible atoms at its edges, at each edge whose paragraph's
 * visible text the selection covers whole (nothing visible before its start in that paragraph,
 * nothing visible after its end). Unchanged otherwise: a selection over part of a paragraph's
 * text keeps its markers where they are.
 */
export function widenedOverInvisible(state: EditorState): TextSelection | null {
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || sel.empty) return null;
  const { doc } = state;
  let { from, to } = sel;
  const $to = doc.resolve(to);
  if ($to.parent.isTextblock) {
    const bounds = visibleBounds($to);
    if (from <= bounds.start && to >= bounds.end) to = $to.end();
  }
  const $from = doc.resolve(from);
  if ($from.parent.isTextblock) {
    const bounds = visibleBounds($from);
    if (from <= bounds.start && to >= bounds.end) from = $from.start();
  }
  if (from === sel.from && to === sel.to) return null;
  const forward = sel.anchor <= sel.head;
  return TextSelection.create(doc, forward ? from : to, forward ? to : from);
}

export function invisibleAtomsPlugin(): Plugin {
  return new Plugin({
    key: invisibleAtomsKey,
    appendTransaction(trs: readonly Transaction[], _old: EditorState, state: EditorState): Transaction | null {
      if (!trs.some((tr) => tr.selectionSet)) return null;
      const widened = widenedOverInvisible(state);
      return widened ? state.tr.setSelection(widened) : null;
    },
    props: {
      handleKeyDown(view, event) {
        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return false;
        if (event.altKey || event.ctrlKey || event.metaKey) return false;
        const { state } = view;
        const { $head, empty } = state.selection;
        if (!empty && !event.shiftKey) return false;
        const dir: -1 | 1 = event.key === 'ArrowRight' ? 1 : -1;
        const neighbour = dir > 0 ? $head.nodeAfter : $head.nodeBefore;
        if (!isInvisible(neighbour)) return false;
        const target = skipInvisible(state, $head.pos, dir);
        if (target === $head.pos) return false;
        const sel = event.shiftKey ? TextSelection.create(state.doc, state.selection.anchor, target) : TextSelection.create(state.doc, target);
        view.dispatch(state.tr.setSelection(sel).scrollIntoView());
        return true;
      },
    },
  });
}

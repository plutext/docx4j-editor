// SPDX-License-Identifier: GPL-3.0-or-later
//
// Down and Up into an empty paragraph (ED-005 section 12.80 item 8). With the formatting marks shown, an empty
// paragraph's trailing break is hidden and its pilcrow is a pseudo-element, so the browser's own vertical caret
// motion finds nothing to land on there and skips the paragraph. When the caret is on the last line of its block
// (or the first, going up) and the next block in that direction is an empty paragraph, the caret is put there,
// before the pilcrow, as it lands without the marks; anything else stays the browser's.
import { Plugin, TextSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';

/** The position inside the empty textblock that starts (`dir` 1) or ends (`dir` -1) the block `node` at `pos`, if that is what it holds. */
function emptyTextblockEdge(node: PMNode, pos: number, dir: 1 | -1): number | undefined {
  let n = node;
  let p = pos;
  for (;;) {
    if (n.isTextblock) return n.content.size === 0 ? p + 1 : undefined;
    if (n.isAtom || n.childCount === 0 || n.type.name === 'table') return undefined;
    if (dir === 1) { n = n.firstChild!; p += 1; }
    else { const last = n.lastChild!; p += n.nodeSize - 1 - last.nodeSize; n = last; }
  }
}

function move(dir: 1 | -1) {
  return (state: EditorState, dispatch?: (tr: ReturnType<EditorState['tr']['setSelection']>) => void, view?: EditorView): boolean => {
    const { selection } = state;
    if (!(selection instanceof TextSelection) || !selection.empty || !view) return false;
    if (!view.endOfTextblock(dir === 1 ? 'down' : 'up')) return false;
    const $pos = selection.$from;
    let depth = $pos.depth;
    while (depth > 0 && !$pos.node(depth).isTextblock) depth--;
    if (depth === 0) return false;
    const $edge = state.doc.resolve(dir === 1 ? $pos.after(depth) : $pos.before(depth));
    const neighbour = dir === 1 ? $edge.nodeAfter : $edge.nodeBefore;
    if (!neighbour) return false;
    const at = dir === 1 ? $edge.pos : $edge.pos - neighbour.nodeSize;
    const target = emptyTextblockEdge(neighbour, at, dir);
    if (target === undefined) return false;
    dispatch?.(state.tr.setSelection(TextSelection.create(state.doc, target)).scrollIntoView());
    return true;
  };
}

export function verticalMotionPlugin(): Plugin {
  const down = move(1);
  const up = move(-1);
  return new Plugin({
    props: {
      handleKeyDown(view, event) {
        if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return false;
        if (event.key === 'ArrowDown') return down(view.state, view.dispatch, view);
        if (event.key === 'ArrowUp') return up(view.state, view.dispatch, view);
        return false;
      },
    },
  });
}

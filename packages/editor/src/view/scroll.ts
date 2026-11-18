// SPDX-License-Identifier: GPL-3.0-or-later
//
// A transaction's `scrollIntoView()` while the focus is elsewhere. ProseMirror scrolls to the DOM
// selection, and writes the state's selection into the DOM only while the editor has the focus
// (prosemirror-view: `selectionToDOM` through `editorOwnsSelection`, then `scrollToSelection`
// returning when the DOM selection's focus node is outside the editor). So the find bar's Next, with
// its query field holding the focus, selected the hit without scrolling to it, until a click in the
// document had put the DOM selection inside the editor. Every surface's dispatch calls this after a
// transaction that asked to scroll: when ProseMirror could not, the element the selection's head is
// drawn in is brought into view.
import { NodeSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export function scrollSelectionWhileUnfocused(view: EditorView): void {
  const focus = view.dom.ownerDocument.getSelection()?.focusNode;
  if (focus && view.dom.contains(focus.nodeType === 1 ? focus : focus.parentNode)) return; // ProseMirror's own scroll
  const sel = view.state.selection;
  let el: Element | null = null;
  try {
    if (sel instanceof NodeSelection) {
      const dom = view.nodeDOM(sel.from);
      el = dom && dom.nodeType === 1 ? (dom as Element) : null;
    } else {
      const { node } = view.domAtPos(sel.head, -1);
      el = node.nodeType === 3 ? node.parentElement : (node as Element);
    }
  } catch {
    return;
  }
  if (!el || typeof el.scrollIntoView !== 'function' || el.getClientRects().length === 0) return;
  el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/** The pixels left above a block brought to the top of its scroll container. */
const TOP_MARGIN = 16;

/**
 * Scroll the block holding `pos` to the top of its scroll container, as a jump to a place does (a TOC entry's
 * Ctrl+click: ED-005 section 12.94). A transaction's `scrollIntoView()` scrolls the least that shows the caret,
 * which lands a target below the fold at the screen's bottom, its content out of view. The container is the
 * nearest scrolling ancestor; without one, the browser's `scrollIntoView` at the block's start.
 */
export function scrollBlockToTop(view: EditorView, pos: number): void {
  let el: Element | null = null;
  try {
    const $pos = view.state.doc.resolve(pos);
    const depth = $pos.depth;
    const dom = depth > 0 ? view.nodeDOM($pos.before(depth)) : view.nodeDOM(pos);
    if (dom && dom.nodeType === 1) el = dom as Element;
    else {
      const { node } = view.domAtPos(pos, -1);
      el = node.nodeType === 3 ? node.parentElement : (node as Element);
    }
  } catch {
    return;
  }
  if (!el || el.getClientRects().length === 0) return;
  const container = scrollingAncestor(el);
  if (!container) { el.scrollIntoView({ block: 'start', inline: 'nearest' }); return; }
  const top = el.getBoundingClientRect().top - container.getBoundingClientRect().top;
  container.scrollTop += top - TOP_MARGIN;
}

/** The nearest ancestor that scrolls, if any does yet. */
export function scrollingAncestor(el: Element): Element | null {
  const win = el.ownerDocument.defaultView;
  for (let node = el.parentElement; node && win; node = node.parentElement) {
    const { overflowY } = win.getComputedStyle(node);
    if ((overflowY === 'auto' || overflowY === 'scroll') && node.scrollHeight > node.clientHeight) return node;
  }
  return null;
}


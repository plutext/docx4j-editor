// SPDX-License-Identifier: GPL-3.0-or-later
//
// The three views of tracked changes (ED-005 section 6.4): All markup, As accepted, Original. The
// view is this plugin's state, set by a transaction's meta, and a class on the body view's root
// (`markup-accepted`, `markup-original`), from which `editor.css` hides and shows over the classes
// the node views set: switching costs one style recalculation. Original is read-only. A document
// always opens in All markup.
//
// Hidden content next to the caret is what the rest of this file is for (section 6.4; the spike,
// section 12.9). ProseMirror maps a position at a boundary to the DOM node before it, so a caret at
// the end of a hidden run, or inside one, lies in hidden DOM and is not drawn, is not scrolled to,
// and makes Backspace and Delete misbehave (Chromium lands the selection elsewhere, Firefox's key
// does nothing). So in a view that hides:
//
// - a collapsed caret at or inside a hidden run moves to whichever edge of the run maps to a drawn
//   place (both edges are one place on screen): the run's start when something visible precedes
//   it in the paragraph, else its end, where a zero-width anchor widget (outside every mark) gives
//   the DOM a point after it; a caret in a hidden paragraph or row moves to the nearest visible
//   textblock in the direction it was going;
// - Backspace and Delete next to hidden content cross it and delete the visible character, atom
//   or paragraph boundary beyond, through the keymap for a boundary;
// - a peer's caret is drawn at the hidden run's end, since its widget takes the marks of what
//   follows it (`clampCaret`, used by presence.ts);
// - a navigation into hidden content (a panel entry, a find hit) switches to All markup
//   (`revealsHidden`), which the callers do.
import { Plugin, PluginKey, TextSelection, Selection, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode, ResolvedPos } from 'prosemirror-model';
import { graphemeBoundary, type Revision } from '@docx4j/editor-model';
import { isInvisible } from './plugins/invisibleAtoms.js';

export type MarkupView = 'all' | 'accepted' | 'original';

export const markupKey = new PluginKey<MarkupState>('docx4j-markup');

interface MarkupState {
  view: MarkupView;
  anchors: DecorationSet;
}

/** The side of a revision a view hides: As accepted hides deletions, Original insertions. */
function hiddenSide(view: MarkupView): 'deletion' | 'insertion' | null {
  return view === 'accepted' ? 'deletion' : view === 'original' ? 'insertion' : null;
}

/** Zero-width items (bookmarks, field codes, ...): drawn as nothing, neither visible content nor hidden. */
const transparent = (node: PMNode): boolean => isInvisible(node);

/** Whether an inline node is hidden in a view: text with its hidden mark, an atom with its hidden side. */
export function isHiddenInline(node: PMNode, view: MarkupView): boolean {
  const side = hiddenSide(view);
  if (!side) return false;
  if (node.isText) return node.marks.some((m) => m.type.name === side);
  const r = node.attrs['revision'] as Revision | null | undefined;
  return !!r?.[side];
}

/** Whether every inline item of a paragraph (zero-width ones aside) carries a side of a revision: nothing of it is left once that side goes. An empty paragraph counts. */
export function allOn(paragraph: PMNode, side: 'deletion' | 'insertion'): boolean {
  const view: MarkupView = side === 'deletion' ? 'accepted' : 'original';
  let all = true;
  paragraph.forEach((child) => { if (all && !transparent(child) && !isHiddenInline(child, view)) all = false; });
  return all;
}

/** Whether a block is hidden whole in a view: a row on the hidden side, a paragraph whose mark is and whose content is all on it. */
export function isHiddenBlock(node: PMNode, view: MarkupView): boolean {
  const side = hiddenSide(view);
  if (!side) return false;
  if (node.type.name === 'table_row') return !!(node.attrs['revision'] as Revision | null)?.[side];
  if (node.type.name === 'paragraph') return !!(node.attrs['markRevision'] as Revision | null)?.[side] && allOn(node, side);
  return false;
}

/** The outermost hidden block around a position, if any. */
function hiddenAncestor($pos: ResolvedPos, view: MarkupView): { before: number; after: number } | undefined {
  for (let d = 1; d <= $pos.depth; d++) if (isHiddenBlock($pos.node(d), view)) return { before: $pos.before(d), after: $pos.after(d) };
  return undefined;
}

interface Kid { node: PMNode; from: number; to: number }

function kidsOf($pos: ResolvedPos): Kid[] {
  const start = $pos.start();
  const out: Kid[] = [];
  $pos.parent.forEach((node, offset) => out.push({ node, from: start + offset, to: start + offset + node.nodeSize }));
  return out;
}

/** The end of a paragraph's leading hidden run (zero-width items within it), where an anchor widget goes; undefined when its first content is visible or it has none hidden. */
function leadingHiddenEnd(kids: Kid[], view: MarkupView): number | undefined {
  let end: number | undefined;
  for (const k of kids) {
    if (transparent(k.node)) continue;
    if (!isHiddenInline(k.node, view)) break;
    end = k.to;
  }
  return end;
}

/**
 * Where a collapsed caret at `pos` is drawn in a view (section 6.4): `pos` itself when the DOM
 * point it maps to is visible; else an edge of the hidden run it is at or in, the run's start
 * when something visible precedes it, else its end (the anchor); in a hidden paragraph or row,
 * the nearest visible textblock in `dir`.
 */
export function visibleCaret(doc: PMNode, pos: number, view: MarkupView, dir: 1 | -1 = 1): number {
  if (view === 'all') return pos;
  const $pos = doc.resolve(pos);
  if (!$pos.parent.inlineContent) return pos;
  const block = hiddenAncestor($pos, view);
  if (block) return nearestVisible(doc, block, view, dir) ?? pos;
  const kids = kidsOf($pos);
  const hidden = (k: Kid): boolean => isHiddenInline(k.node, view);
  let j = 0;
  while (j < kids.length && kids[j]!.to <= pos) j++;
  let r: number;
  if (j < kids.length && kids[j]!.from < pos) {
    // Inside a text node.
    if (!hidden(kids[j]!)) return pos;
    r = j;
  } else {
    let b = j - 1;
    while (b >= 0 && transparent(kids[b]!.node)) b--;
    if (b >= 0 && !hidden(kids[b]!)) return pos;
    if (b < 0) {
      // At the paragraph's start: drawn unless the first content is hidden.
      let f = j;
      while (f < kids.length && transparent(kids[f]!.node)) f++;
      if (f === kids.length || !hidden(kids[f]!)) return pos;
      return leadingHiddenEnd(kids, view) ?? pos;
    }
    r = b;
  }
  // The run holding kid r: back to its first hidden kid, then what precedes it.
  let s = r;
  while (s - 1 >= 0 && (hidden(kids[s - 1]!) || transparent(kids[s - 1]!.node))) s--;
  while (s < r && transparent(kids[s]!.node)) s++;
  if (s > 0) return kids[s]!.from;
  return leadingHiddenEnd(kids, view) ?? pos;
}

/** The nearest caret outside a hidden block, `dir` first, then the other way. */
function nearestVisible(doc: PMNode, block: { before: number; after: number }, view: MarkupView, dir: 1 | -1): number | undefined {
  for (const d of [dir, -dir as 1 | -1]) {
    let at = d > 0 ? block.after : block.before;
    for (let guard = 0; guard < 200; guard++) {
      if (at < 0 || at > doc.content.size) break;
      const found = Selection.findFrom(doc.resolve(at), d, true);
      if (!found) break;
      const $h = found.$head;
      const inner = hiddenAncestor($h, view);
      if (!inner) return visibleCaret(doc, $h.pos, view, d);
      at = d > 0 ? inner.after : inner.before;
    }
  }
  return undefined;
}

/**
 * Where a peer's caret is drawn in this view (presence.ts). Not where a local caret goes: the
 * peer's caret is y-prosemirror's widget (`side` 10), which takes the marks of what follows it,
 * so at a hidden run's start it is drawn inside the hidden element and is not seen (the spike,
 * section 12.9). So the end of the hidden run the caret is at or in; in a hidden paragraph or row,
 * the nearest visible textblock.
 */
export function clampCaret(state: EditorState, pos: number): number {
  const view = markupKey.getState(state)?.view ?? 'all';
  const doc = state.doc;
  if (view === 'all' || pos < 0 || pos > doc.content.size) return pos;
  const $pos = doc.resolve(pos);
  if (!$pos.parent.inlineContent) return pos;
  const block = hiddenAncestor($pos, view);
  if (block) return nearestVisible(doc, block, view, 1) ?? pos;
  let at = pos;
  // Inside a hidden text node: from its end on.
  if ($pos.textOffset > 0 && $pos.parent.child($pos.index()).isText && isHiddenInline($pos.parent.child($pos.index()), view)) at = pos - $pos.textOffset + $pos.parent.child($pos.index()).nodeSize;
  // Over what follows while it is hidden or zero-width: the widget then goes before visible content, or at the paragraph's end.
  for (let guard = 0; guard < 10_000; guard++) {
    const n = doc.resolve(at).nodeAfter;
    if (!n || !(isHiddenInline(n, view) || transparent(n))) break;
    at += n.nodeSize;
  }
  return at;
}

/** Whether a range, or a caret, lies in content the current view hides: a navigation there switches to All markup. */
export function revealsHidden(state: EditorState, from: number, to: number): boolean {
  const view = markupKey.getState(state)?.view ?? 'all';
  if (view === 'all') return false;
  const doc = state.doc;
  if (hiddenAncestor(doc.resolve(from), view)) return true;
  if (from === to) {
    const $pos = doc.resolve(from);
    return $pos.parent.inlineContent && visibleCaret(doc, from, view) !== from;
  }
  let visible = false;
  doc.nodesBetween(from, to, (node) => {
    if (visible) return false;
    if (isHiddenBlock(node, view)) return false;
    if (node.isInline && !transparent(node) && !isHiddenInline(node, view)) visible = true;
    return !visible;
  });
  return !visible;
}

/** The anchors of a view: a zero-width widget at the end of each paragraph's leading hidden run. Memoised per top-level block and view. */
const anchorMemo = new WeakMap<PMNode, Partial<Record<MarkupView, number[]>>>();

function anchorsIn(block: PMNode, view: MarkupView): number[] {
  let memo = anchorMemo.get(block);
  if (!memo) { memo = {}; anchorMemo.set(block, memo); }
  const known = memo[view];
  if (known) return known;
  const out: number[] = [];
  const visit = (node: PMNode, pos: number): void => {
    if (node.type.name === 'paragraph') {
      if (isHiddenBlock(node, view)) return;
      const kids: Kid[] = [];
      node.forEach((child, offset) => kids.push({ node: child, from: pos + 1 + offset, to: pos + 1 + offset + child.nodeSize }));
      const end = leadingHiddenEnd(kids, view);
      if (end !== undefined) out.push(end);
      return;
    }
    if (isHiddenBlock(node, view) || node.isLeaf || node.inlineContent) return;
    node.forEach((child, offset) => visit(child, pos + 1 + offset));
  };
  visit(block, 0);
  memo[view] = out;
  return out;
}

function anchorWidget(): HTMLElement {
  const el = document.createElement('span');
  el.className = 'rev-anchor';
  el.textContent = String.fromCharCode(0x200b); // a zero-width space: a DOM point after it that is drawn
  return el;
}

function anchorsOf(doc: PMNode, view: MarkupView): DecorationSet {
  if (view === 'all') return DecorationSet.empty;
  const decos: Decoration[] = [];
  doc.forEach((block, offset) => {
    for (const rel of anchorsIn(block, view)) decos.push(Decoration.widget(offset + rel, anchorWidget, { side: -1, marks: [], key: 'rev-anchor' }));
  });
  return DecorationSet.create(doc, decos);
}

/**
 * Where a walk from a textblock's edge goes on to when the paragraph mark there is hidden (a
 * deleted mark in As accepted, an inserted one in Original): the end of the paragraph before, or
 * the start of the one after, since the view draws the two as one. Undefined at a mark it shows.
 */
function acrossMark($p: ResolvedPos, dir: -1 | 1, view: MarkupView): number | undefined {
  const side = hiddenSide(view);
  if (!side || $p.parent.type.name !== 'paragraph' || $p.depth < 1) return undefined;
  const container = $p.node($p.depth - 1);
  const index = $p.index($p.depth - 1);
  if (dir < 0) {
    if ($p.parentOffset !== 0 || index === 0) return undefined;
    const prev = container.child(index - 1);
    return prev.type.name === 'paragraph' && (prev.attrs['markRevision'] as Revision | null)?.[side] ? $p.before() - 1 : undefined;
  }
  if ($p.parentOffset !== $p.parent.content.size || index + 1 >= container.childCount) return undefined;
  const next = container.child(index + 1);
  return next.type.name === 'paragraph' && ($p.parent.attrs['markRevision'] as Revision | null)?.[side] ? $p.after() + 1 : undefined;
}

/**
 * Backspace (`dir` -1) or Delete (1) next to content the view hides: the hidden content crossed,
 * a hidden paragraph mark included, then the visible grapheme or atom beyond it deleted (the view's
 * dispatch tracks it as any edit), or, at a paragraph's edge the view shows, the caret put there
 * for the keymap to join as it does. Elsewhere the browser's own deletion is right, and it is left.
 */
function deleteAcross(view: EditorView, dir: -1 | 1): boolean {
  const state = view.state;
  const mode = markupKey.getState(state)?.view ?? 'all';
  if (mode === 'all') return false;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty || !sel.$head.parent.inlineContent) return false;
  const doc = state.doc;
  /** From `pos` in `d`: over zero-width items, and, when `hidden`, over hidden content and hidden marks; where it stops, and whether it crossed anything hidden. */
  const walk = (from: number, d: -1 | 1, hidden: boolean): { pos: number; crossed: boolean } => {
    let pos = from;
    let crossed = false;
    for (let guard = 0; guard < 10_000; guard++) {
      const $p = doc.resolve(pos);
      const n = d < 0 ? $p.nodeBefore : $p.nodeAfter;
      if (n && transparent(n)) { pos += d * n.nodeSize; continue; }
      if (!hidden) break;
      if (n && isHiddenInline(n, mode)) { pos += d * n.nodeSize; crossed = true; continue; }
      const across = n ? undefined : acrossMark($p, d, mode);
      if (across === undefined) break;
      pos = across;
      crossed = true;
    }
    return { pos, crossed };
  };
  // Only where hidden content is next to the caret, on either side (zero-width items aside).
  const near = (d: -1 | 1): boolean => {
    const { pos } = walk(sel.head, d, false);
    const $p = doc.resolve(pos);
    const n = d < 0 ? $p.nodeBefore : $p.nodeAfter;
    return n ? isHiddenInline(n, mode) : acrossMark($p, d, mode) !== undefined;
  };
  if (!near(-1) && !near(1)) return false;
  const { pos } = walk(sel.head, dir, true);
  const $at = doc.resolve(pos);
  const beyond = dir < 0 ? $at.nodeBefore : $at.nodeAfter;
  if (!beyond) {
    // A paragraph's edge the view shows: the keymap joins from there (its list rules included).
    if (pos !== sel.head) view.dispatch(state.tr.setSelection(TextSelection.create(doc, pos)));
    return false;
  }
  let from: number;
  let to: number;
  if (beyond.isText) {
    const text = beyond.text!;
    if (dir < 0) { const cut = graphemeBoundary(text, text.length - 1, 'backward'); from = pos - (text.length - cut); to = pos; }
    else { const cut = graphemeBoundary(text, 1, 'forward'); from = pos; to = pos + cut; }
  } else {
    from = dir < 0 ? pos - beyond.nodeSize : pos;
    to = dir < 0 ? pos : pos + beyond.nodeSize;
  }
  const tr = state.tr.setSelection(TextSelection.create(doc, dir < 0 ? to : from)).delete(from, to);
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** The meta a transaction carries to change the view: `{ view }`. */
export function setMarkupView(tr: Transaction, view: MarkupView): Transaction {
  return tr.setMeta(markupKey, { view });
}

export function markupPlugin(): Plugin<MarkupState> {
  return new Plugin<MarkupState>({
    key: markupKey,
    state: {
      init: () => ({ view: 'all', anchors: DecorationSet.empty }),
      apply(tr, value, _old, next) {
        const meta = tr.getMeta(markupKey) as { view: MarkupView } | undefined;
        const view = meta?.view ?? value.view;
        if (view === 'all') return value.view === 'all' ? value : { view, anchors: DecorationSet.empty };
        if (meta || tr.docChanged) return { view, anchors: anchorsOf(next.doc, view) };
        return value;
      },
    },
    props: {
      attributes: (state): Record<string, string> => {
        const view = markupKey.getState(state)?.view ?? 'all';
        return view === 'all' ? {} : { class: `markup-${view}` };
      },
      editable: (state) => markupKey.getState(state)?.view !== 'original',
      // ProseMirror scrolls to `coordsAtPos(head, 1)`, the side after the caret, which at a hidden
      // run's start is the hidden text: a rectangle at 0,0 it takes as already in view. While a view
      // hides content, the element the caret's DOM point is in is brought into view instead.
      handleScrollToSelection(view) {
        if ((markupKey.getState(view.state)?.view ?? 'all') === 'all') return false;
        const sel = view.state.selection;
        if (!(sel instanceof TextSelection)) return false;
        const { node } = view.domAtPos(sel.head, -1);
        const el = node.nodeType === 3 ? node.parentElement : (node as Element);
        if (!el || typeof el.scrollIntoView !== 'function' || el.getClientRects().length === 0) return false;
        el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        return true;
      },
      decorations: (state) => markupKey.getState(state)?.anchors,
      handleKeyDown(view, event) {
        if (event.ctrlKey || event.metaKey || event.altKey) return false;
        if (event.key === 'Backspace') return deleteAcross(view, -1);
        if (event.key === 'Delete') return deleteAcross(view, 1);
        return false;
      },
    },
    appendTransaction(trs, oldState, newState) {
      const view = markupKey.getState(newState)?.view ?? 'all';
      if (view === 'all') return null;
      const sel = newState.selection;
      if (!(sel instanceof TextSelection) || !sel.empty) return null;
      if (!trs.some((tr) => tr.selectionSet || tr.docChanged || tr.getMeta(markupKey))) return null;
      const dir: 1 | -1 = sel.head >= oldState.selection.head ? 1 : -1;
      const target = visibleCaret(newState.doc, sel.head, view, dir);
      if (target === sel.head) return null;
      return newState.tr.setSelection(TextSelection.create(newState.doc, target));
    },
  });
}

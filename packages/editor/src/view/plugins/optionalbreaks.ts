// SPDX-License-Identifier: GPL-3.0-or-later
//
// Optional breaks drawn (ED-007 section 8.1 item 13; Jason, 2026-10-15): a zero-width space, U+200B, is a character
// a line may break at and nothing is drawn for, in Word and here. With the formatting marks shown, Word draws a small
// box for it; so does the view, through an inline decoration on the character (`editor.css`, `.optional-break`), the
// one formatting mark a stylesheet rule cannot place, since it is a character inside a text node. The decorations are
// memoised per top-level block (the characters are rare, the walk is not), and a click on the mark selects the
// character, so that Delete or Backspace removes it, as the keys do at a caret beside it.
import { Plugin, PluginKey, TextSelection } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { OPTIONAL_BREAK } from '@docx4j/editor-model';

export const OPTIONAL_BREAK_CLASS = 'optional-break';

const key = new PluginKey<DecorationSet>('optionalBreaks');

/** The offsets of every optional break inside a block, relative to the block's position; cached by the block node. */
const cache = new WeakMap<PMNode, number[]>();

function offsetsIn(block: PMNode): number[] {
  let found = cache.get(block);
  if (found) return found;
  const out: number[] = [];
  block.descendants((node, pos) => {
    if (!node.isText) return true;
    const text = node.text ?? '';
    for (let i = text.indexOf(OPTIONAL_BREAK); i >= 0; i = text.indexOf(OPTIONAL_BREAK, i + 1)) out.push(pos + 1 + i);
    return false;
  });
  found = out;
  cache.set(block, found);
  return found;
}

/** Every optional break in the document as an inline decoration; the empty set when there is none. */
export function optionalBreakDecorations(doc: PMNode): DecorationSet {
  const decorations: Decoration[] = [];
  doc.forEach((block, pos) => {
    for (const offset of offsetsIn(block)) decorations.push(Decoration.inline(pos + offset, pos + offset + 1, { class: OPTIONAL_BREAK_CLASS }));
  });
  return decorations.length === 0 ? DecorationSet.empty : DecorationSet.create(doc, decorations);
}

/** The optional break the selection is exactly over, if any: its mark is drawn highlighted. */
function selectedBreak(doc: PMNode, from: number, to: number): DecorationSet | undefined {
  if (to - from !== 1 || doc.textBetween(from, to) !== OPTIONAL_BREAK) return undefined;
  return DecorationSet.create(doc, [Decoration.inline(from, to, { class: `${OPTIONAL_BREAK_CLASS} selected` })]);
}

export function optionalBreaksPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: (_config, state) => optionalBreakDecorations(state.doc),
      apply: (tr, set, _old, state) => (tr.docChanged ? optionalBreakDecorations(state.doc) : set),
    },
    props: {
      decorations: (state) => {
        const base = key.getState(state) ?? DecorationSet.empty;
        const selected = selectedBreak(state.doc, state.selection.from, state.selection.to);
        return selected ? selected.add(state.doc, base.find()) : base;
      },
      // A click on the mark selects the character under it, so that Delete removes it.
      handleClick: (view, _pos, event) => {
        const target = event.target as Element | null;
        const el = target?.closest?.(`.${OPTIONAL_BREAK_CLASS}`);
        if (!el || !view.dom.contains(el)) return false;
        const at = view.posAtDOM(el, 0);
        if (view.state.doc.textBetween(at, at + 1) !== OPTIONAL_BREAK) return false;
        view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at + 1)));
        return true;
      },
    },
  });
}

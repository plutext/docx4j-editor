// SPDX-License-Identifier: GPL-3.0-or-later
//
// The caret and the view kept across a reload, and carried into the disk version's tab (ED-011 section 3.5). What is
// kept is plain data, so it outlives the session and crosses to another tab: the paragraphs' text, the caret as a
// place in it, and the anchor the view is held by, with the height on screen it stood at.
//
// The anchor is the caret while the caret is on screen, and otherwise the first line in view: a caret left behind, or
// never placed, is not where a person is reading, so the view is kept and the caret comes back without scrolling to it.
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { graphemeBoundary, paragraphStrings, positionOfPlace, relocator, textPlaceOf, type Landing, type TextPlace } from '@docx4j/editor-model';
import type { DocxEditorView } from './DocxEditorView.js';
import type { MarkupView } from './markup.js';
import { scrollingAncestor } from './scroll.js';

export interface KeptPlace {
  paragraphs: string[];
  caret: TextPlace;
  anchor: TextPlace;
  /** The anchor line's top, in pixels below the scroll container's top edge. */
  top: number;
  markup: MarkupView;
}

/** What became of the caret's text. */
export interface PlaceRestored {
  found: Landing['found'];
  /** The opening words of the paragraph the caret was in, for saying what went. */
  opening: string;
  /** Stop holding the view. */
  stop: () => void;
}

/** The words a status line quotes. */
const QUOTED = 40;

/** A paragraph's opening words on one line. The cut is measured on a short slice: a grapheme walk is quadratic in its text. */
function openingOf(text: string): string {
  const line = text.replace(/\s+/g, ' ').trim();
  return line.length <= QUOTED ? line : `${line.slice(0, graphemeBoundary(line.slice(0, QUOTED * 2), QUOTED)).trimEnd()}…`;
}

function anchorOf(view: EditorView, head: number): { pos: number; top: number } {
  const scroller = scrollingAncestor(view.dom);
  if (!scroller) return { pos: head, top: 0 };
  const box = scroller.getBoundingClientRect();
  const caret = view.coordsAtPos(head);
  if (caret.top >= box.top && caret.bottom <= box.bottom) return { pos: head, top: caret.top - box.top };
  const page = view.dom.getBoundingClientRect();
  const pos = view.posAtCoords({ left: page.left + page.width / 2, top: Math.max(box.top, page.top) + 1 })?.pos ?? head;
  return { pos, top: view.coordsAtPos(pos).top - box.top };
}

/** The body's place, read before the view goes: its own caret, wherever the focus is, as a note's caret is in no version of the body. */
export function keepPlace(editor: DocxEditorView): KeptPlace {
  const { view } = editor;
  const { doc, selection } = view.state;
  // A selection comes back as a caret. A cell selection's `from` is its head cell's, a row's last, so its ranges say where it starts.
  const at = selection instanceof TextSelection ? selection.head : Math.min(...selection.ranges.map((r) => r.$from.pos));
  const { pos, top } = anchorOf(view, at);
  return { paragraphs: paragraphStrings(doc), caret: textPlaceOf(doc, at), anchor: textPlaceOf(doc, pos), top, markup: editor.markupView };
}

/**
 * Keep `pos` at `top` while the page settles and until the person moves: embedded fonts, a Markdown import's table fit
 * and the measured tab stops all reflow the page after the first paint. Only a change of the page's size acts.
 */
function holdView(view: EditorView, pos: number, top: number): () => void {
  const align = (): void => {
    const scroller = view.isDestroyed ? null : scrollingAncestor(view.dom);
    if (!scroller) return;
    const at = Math.min(pos, view.state.doc.content.size);
    scroller.scrollTop += view.coordsAtPos(at).top - scroller.getBoundingClientRect().top - top;
  };
  align();
  const win = view.dom.ownerDocument.defaultView ?? window;
  const inputs = ['wheel', 'keydown', 'pointerdown', 'touchstart'] as const;
  const observer = new ResizeObserver(align);
  const stop = (): void => {
    observer.disconnect();
    for (const input of inputs) win.removeEventListener(input, stop, true);
  };
  observer.observe(view.dom);
  for (const input of inputs) win.addEventListener(input, stop, { capture: true, passive: true });
  return stop;
}

/** The kept place found again in the document now open: the caret put back, the view held, nothing recorded as an edit. */
export function restorePlace(editor: DocxEditorView, kept: KeptPlace): PlaceRestored {
  const { view } = editor;
  // First: a view hides text, and the heights below are read from the view that will stay.
  editor.markupView = kept.markup;
  const { doc } = view.state;
  const land = relocator(kept.paragraphs, paragraphStrings(doc));
  const caret = land(kept.caret);
  view.dispatch(view.state.tr.setSelection(TextSelection.near(doc.resolve(positionOfPlace(doc, caret)))));
  const opening = openingOf(kept.paragraphs[kept.caret.paragraph] ?? '');
  // Nothing in common: no place to keep, and a fresh page is at its start.
  if (caret.found === 'unrelated') return { found: caret.found, opening, stop: () => undefined };
  return { found: caret.found, opening, stop: holdView(view, positionOfPlace(doc, land(kept.anchor)), kept.top) };
}

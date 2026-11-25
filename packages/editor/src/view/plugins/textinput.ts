// SPDX-License-Identifier: GPL-3.0-or-later
//
// Text input narrowed to what changed (ED-013 section 9, S5 defect 2), and Chrome on Android's Enter given to the keymaps
// (S5's second part; `handleEnterInput` below).
//
// ProseMirror reads typing from the DOM. During a composition in Chrome, from its second update on, the view
// description of a mark around the composed text is dirty, so ProseMirror re-parses the mark's element through the
// schema, where no mark has a parse rule (paste has its own converter). The parsed text lacks the mark, and the
// change ProseMirror finds is the whole text node. Its default, `insertText` over that range, takes the marks back,
// but the caret is mapped through the whole replacement, to its end. ProseMirror's rule for Chrome's composition
// selection then distrusts the browser's caret, which is right, because it lies one before that end whenever one
// character follows it in the node: a sentence's final period, or a space. Samsung Keyboard ends the composition
// at the wrong caret, and what is typed next lands after the period.
//
// Here such a change is narrowed to the text that differs (its common prefix and suffix with the text it replaces
// dropped), as ProseMirror's own diff would have found it. The caret is ProseMirror's own (the browser's), but during
// a composition, where that rule may have replaced the browser's with the mapped one, it is kept only when it lies in
// what was typed; else the old caret is mapped through the narrowed change. Outside a composition (a suggestion
// picked: the word deleted, the new one and a space inserted) the browser's caret, after the space, is kept.
import { Plugin, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

const isHighSurrogate = (code: number): boolean => code >= 0xd800 && code <= 0xdbff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc00 && code <= 0xdfff;

/** The common prefix and suffix of two strings, never splitting a surrogate pair, and never overlapping in either. */
export function commonEnds(a: string, b: string): { prefix: number; suffix: number } {
  const max = Math.min(a.length, b.length);
  let prefix = 0;
  while (prefix < max && a.charCodeAt(prefix) === b.charCodeAt(prefix)) prefix++;
  if (prefix > 0 && prefix < a.length && isLowSurrogate(a.charCodeAt(prefix))) prefix--;
  let suffix = 0;
  while (suffix < max - prefix && a.charCodeAt(a.length - 1 - suffix) === b.charCodeAt(b.length - 1 - suffix)) suffix++;
  if (suffix > 0 && a.length - suffix > 0 && isHighSurrogate(a.charCodeAt(a.length - suffix - 1))) suffix--;
  return { prefix, suffix };
}

/**
 * ProseMirror's text input (`handleTextInput`), narrowed when the range it reports is wider than the text that
 * changed. False, and ProseMirror's default, when it is narrow already or holds anything but text.
 */
export function narrowTextInput(view: EditorView, from: number, to: number, text: string, deflt: () => Transaction): boolean {
  // UPSTREAM-27: ProseMirror reports a marked run's whole text node from a composition's second update in Chrome.
  if (to <= from) return false;
  const state = view.state;
  const $from = state.doc.resolve(from);
  if (!$from.sameParent(state.doc.resolve(to)) || !$from.parent.inlineContent) return false;
  const old = state.doc.textBetween(from, to);
  // Text alone: an atom in the range has a size but no text.
  if (old.length !== to - from) return false;
  const { prefix, suffix } = commonEnds(old, text);
  if (prefix === 0 && suffix === 0) return false;
  const typed = text.slice(prefix, text.length - suffix);
  const at = from + prefix;
  const end = to - suffix;
  // ProseMirror's own transaction: its composition meta, and the caret it took from the browser or mapped.
  const base = deflt();
  const tr = state.tr;
  if (typed) tr.insertText(typed, at, end);
  else tr.delete(at, end);
  const composition = base.getMeta('composition') as unknown;
  if (composition !== undefined) tr.setMeta('composition', composition);
  // The two documents hold the same text, so a position in one is the same place in the other.
  const reported = base.selection;
  const inRange = (pos: number): boolean => pos >= 0 && pos <= tr.doc.content.size;
  if (reported instanceof TextSelection && inRange(reported.anchor) && inRange(reported.head)) {
    const typedOver = reported.empty && reported.head > at && reported.head <= at + typed.length;
    if (!view.composing || typedOver) tr.setSelection(TextSelection.create(tr.doc, reported.anchor, reported.head));
  }
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** Chrome on Android, as ProseMirror tells it (`android && chrome`): where ProseMirror leaves the Enter key to the browser. */
export function isChromeAndroid(agent: string = typeof navigator === 'undefined' ? '' : navigator.userAgent): boolean {
  return /Android \d/.test(agent) && /Chrome\/\d/.test(agent);
}

/** A keydown as ProseMirror's own `keyEvent` makes one, for the keymaps' `handleKeyDown`. */
function enterKey(shiftKey: boolean): KeyboardEvent {
  const event = document.createEvent('Event');
  event.initEvent('keydown', true, true);
  Object.assign(event, { keyCode: 13, key: 'Enter', code: 'Enter', shiftKey });
  return event as KeyboardEvent;
}

/**
 * Enter as Chrome on Android delivers it (ED-013 section 9, S5): the keyboard's Enter is a keydown that ProseMirror leaves to
 * the browser there, then the browser's own split (`beforeinput` of `insertParagraph`), which ProseMirror reads back from the
 * DOM. Unless the change it finds looks like an Enter, it applies the change as parsed, and the editor's marks have no parse
 * rules: the paragraph's content is replaced, its text losing its runs (`runSource`) and the new paragraph its properties
 * (`copiedFrom`), and with tracking on the whole paragraph is recorded as deleted and inserted again. Measured on Jason's
 * phone (Samsung Keyboard and Gboard), where every Enter in a document's runs did so. Here the browser's split is cancelled
 * and the keymaps get the Enter (Shift+Enter for `insertLineBreak`) as ProseMirror hands them one it recognises in a DOM
 * change; one the keymaps decline is left to the browser, as before. Not during a composition, nor anywhere but Chrome on
 * Android: elsewhere ProseMirror takes the key itself (and on iOS a cancelled Enter confuses the keyboard).
 */
export function handleEnterInput(view: EditorView, event: InputEvent, android: boolean = isChromeAndroid()): boolean {
  // UPSTREAM-29: on Chrome for Android ProseMirror leaves Enter to the browser and reads the split back.
  if (!android || !event.cancelable || view.composing || event.isComposing) return false;
  if (event.inputType !== 'insertParagraph' && event.inputType !== 'insertLineBreak') return false;
  const handled = view.someProp('handleKeyDown', (f) => f(view, enterKey(event.inputType === 'insertLineBreak'))) ?? false;
  if (handled) event.preventDefault();
  return handled;
}

/**
 * Every editing surface's: text input narrowed to what changed (`narrowTextInput`), and Chrome on Android's Enter given to
 * the keymaps (`handleEnterInput`).
 */
export function textInputPlugin(): Plugin {
  return new Plugin({ props: { handleTextInput: narrowTextInput, handleDOMEvents: { beforeinput: (view, event) => handleEnterInput(view, event as InputEvent) } } });
}

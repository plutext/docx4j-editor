// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 section 9, S5 defect 2: ProseMirror's text input narrowed to what changed; and S5's second part, Chrome on
// Android's Enter given to the keymaps (`handleEnterInput`). From a composition's second update
// in Chrome, ProseMirror reports a marked run's whole text node as changed and maps the caret to its end, past a final
// period; `narrowTextInput` makes it the characters that differ, with the caret after them. The browser test is
// e2e/composition.spec.ts.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { EditorDocument, stateFor } from '@docx4j/editor-model';
import { commonEnds, narrowTextInput, handleEnterInput, isChromeAndroid } from '../src/view/plugins/textinput.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

describe('commonEnds', () => {
  test('the prefix and suffix two texts share, never overlapping', () => {
    expect(commonEnds('Some texta.', 'Some textab.')).toEqual({ prefix: 10, suffix: 1 });
    expect(commonEnds('aa', 'aaa')).toEqual({ prefix: 2, suffix: 0 });
    expect(commonEnds('dont', "don't")).toEqual({ prefix: 3, suffix: 1 });
    expect(commonEnds('abc', 'xyz')).toEqual({ prefix: 0, suffix: 0 });
  });

  test('never splits a surrogate pair', () => {
    // U+1F600 and U+1F603 share their high surrogate.
    expect(commonEnds('x\u{1F600}y', 'x\u{1F603}y')).toEqual({ prefix: 1, suffix: 1 });
    // U+1F600 and U+1D600 share their low surrogate.
    expect(commonEnds('x\u{1F600}', 'x\u{1D600}')).toEqual({ prefix: 1, suffix: 0 });
  });
});

describe('narrowTextInput', () => {
  /** "Some text." with "a" typed after "text", the state and the paragraph's text node's range. */
  async function composing(): Promise<{ state: EditorState; from: number; to: number }> {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    let from = -1;
    session.doc.descendants((n, pos) => { if (from < 0 && n.isText && n.text === 'Some text.') from = pos; return from < 0; });
    expect(from).toBeGreaterThan(0);
    let state = stateFor(session, { from: from + 9 });
    state = state.apply(state.tr.insertText('a', from + 9));
    return { state, from, to: from + 'Some texta.'.length };
  }

  /** A view that is a state, a dispatch and whether it is composing: all `narrowTextInput` uses. */
  function viewOf(state: EditorState, composing = true): { view: EditorView; dispatched: () => Transaction | undefined } {
    let last: Transaction | undefined;
    return { view: { state, composing, dispatch: (tr: Transaction) => { last = tr; } } as unknown as EditorView, dispatched: () => last };
  }

  /** ProseMirror's default for the whole node, its caret at `caret`, with a composition meta. */
  const whole = (state: EditorState, from: number, to: number, text: string, caret: number) => (): Transaction => {
    const tr = state.tr.insertText(text, from, to);
    return tr.setSelection(TextSelection.create(tr.doc, caret)).setMeta('composition', 7);
  };

  test('the whole node reported: the characters typed inserted, the caret after them, the composition meta kept', async () => {
    const { state, from, to } = await composing();
    const { view, dispatched } = viewOf(state);
    // ProseMirror's mapped caret is past the period: from + "Some textab.".length.
    expect(narrowTextInput(view, from, to, 'Some textab.', whole(state, from, to, 'Some textab.', from + 12))).toBe(true);
    const tr = dispatched()!;
    expect(tr.steps.map((s) => s.toJSON() as { from: number; to: number })).toEqual([expect.objectContaining({ from: from + 10, to: from + 10 })]);
    expect(tr.doc.textBetween(from, from + 12)).toBe('Some textab.');
    expect(tr.selection.head).toBe(from + 11);
    expect(tr.getMeta('composition')).toBe(7);
  });

  test("the browser's caret kept when it is in what was typed", async () => {
    const { state, from, to } = await composing();
    const { view, dispatched } = viewOf(state);
    expect(narrowTextInput(view, from, to, 'Some textab.', whole(state, from, to, 'Some textab.', from + 11))).toBe(true);
    expect(dispatched()!.selection.head).toBe(from + 11);
  });

  test('a deletion narrowed to the character deleted', async () => {
    const { state, from, to } = await composing();
    const { view, dispatched } = viewOf(state);
    expect(narrowTextInput(view, from, to, 'Some text.', whole(state, from, to, 'Some text.', from + 10))).toBe(true);
    const tr = dispatched()!;
    expect(tr.doc.textBetween(from, from + 10)).toBe('Some text.');
    expect(tr.selection.head).toBe(from + 9);
  });

  test("outside a composition (a suggestion picked, a space after it) the browser's caret is kept", async () => {
    const { state, from, to } = await composing();
    const { view, dispatched } = viewOf(state, false);
    // "texta" picked as "text": the a deleted, the browser's caret after the period.
    expect(narrowTextInput(view, from, to, 'Some text.', whole(state, from, to, 'Some text.', from + 10))).toBe(true);
    const tr = dispatched()!;
    expect(tr.doc.textBetween(from, from + 10)).toBe('Some text.');
    expect(tr.selection.head).toBe(from + 10);
  });

  test('left to ProseMirror when the change is narrow already, or empty', async () => {
    const { state, from } = await composing();
    const { view, dispatched } = viewOf(state);
    expect(narrowTextInput(view, from + 10, from + 10, 'b', whole(state, from + 10, from + 10, 'b', from + 11))).toBe(false);
    expect(narrowTextInput(view, from, from + 4, 'Wxyz', whole(state, from, from + 4, 'Wxyz', from + 4))).toBe(false);
    expect(dispatched()).toBeUndefined();
  });
});

describe('handleEnterInput', () => {
  const CHROME_ANDROID = 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Mobile Safari/537.36';

  test('Chrome on Android, as ProseMirror tells it', () => {
    expect(isChromeAndroid(CHROME_ANDROID)).toBe(true);
    expect(isChromeAndroid('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36')).toBe(false);
    expect(isChromeAndroid('Mozilla/5.0 (Android 14; Mobile; rv:140.0) Gecko/140.0 Firefox/140.0')).toBe(false);
  });

  /** A view whose keymaps take Enter or not, recording the keydowns they were handed. */
  function viewTaking(takes: boolean, composing = false): { view: EditorView; keys: KeyboardEvent[] } {
    const keys: KeyboardEvent[] = [];
    const view = { composing, someProp: (name: string, f: (h: (v: EditorView, e: KeyboardEvent) => boolean) => boolean) => (name === 'handleKeyDown' ? f((_v, e) => { keys.push(e); return takes; }) : undefined) } as unknown as EditorView;
    return { view, keys };
  }
  const input = (inputType: string, init: InputEventInit = {}): InputEvent => new InputEvent('beforeinput', { inputType, cancelable: true, ...init });

  test("the browser's paragraph split cancelled and Enter given to the keymaps; a line break as Shift+Enter", () => {
    const { view, keys } = viewTaking(true);
    const split = input('insertParagraph');
    expect(handleEnterInput(view, split, true)).toBe(true);
    expect(split.defaultPrevented).toBe(true);
    const lineBreak = input('insertLineBreak');
    expect(handleEnterInput(view, lineBreak, true)).toBe(true);
    expect(keys.map((k) => [k.type, k.keyCode, k.key, k.shiftKey])).toEqual([['keydown', 13, 'Enter', false], ['keydown', 13, 'Enter', true]]);
  });

  test('an Enter the keymaps decline is left to the browser', () => {
    const { view } = viewTaking(false);
    const split = input('insertParagraph');
    expect(handleEnterInput(view, split, true)).toBe(false);
    expect(split.defaultPrevented).toBe(false);
  });

  test('nothing elsewhere than Chrome on Android, during a composition, for an event that cannot be cancelled, or for other input', () => {
    const { view, keys } = viewTaking(true);
    expect(handleEnterInput(view, input('insertParagraph'), false)).toBe(false);
    expect(handleEnterInput(view, input('insertParagraph', { isComposing: true }), true)).toBe(false);
    expect(handleEnterInput(view, new InputEvent('beforeinput', { inputType: 'insertParagraph', cancelable: false }), true)).toBe(false);
    expect(handleEnterInput(view, input('insertText', { data: 'a' }), true)).toBe(false);
    expect(handleEnterInput(viewTaking(true, true).view, input('insertParagraph'), true)).toBe(false);
    expect(keys).toEqual([]);
  });
});

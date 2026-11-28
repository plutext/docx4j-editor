// SPDX-License-Identifier: GPL-3.0-or-later
//
// The key gate (ED-014 section 4.4, P32) in jsdom, over a real body view and a read-only story view, installed by
// `initHost()` with a fake of VS Code's API as the extension's page installs it. A listener on `window` in the bubble phase
// stands for VS Code's host page, which forwards every key that reaches it to the workbench. jsdom has no `DataTransfer`,
// `ClipboardEvent` or `execCommand`, so the plain paste is the browser suite's (`e2e/host-vscode.spec.ts`).
import { afterEach, beforeAll, describe, expect, test } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { StoryView } from '../src/view/StoryView.js';
import { initHost, resetHost } from '../src/app/host.js';
import { pageKeepsKey } from '../src/app/keygate.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

let session: EditorDocument;
let body: DocxEditorView;
let readOnly: HTMLElement;
let stepped = 0;

beforeAll(async () => {
  session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  // The application's shortcuts as App passes them, two of them: the next change (Alt+PageDown) and Insert Link (Ctrl+K).
  body = new DocxEditorView(container, session, { headersAndFooters: false, shortcuts: { 'Alt-PageDown': () => { stepped++; }, 'Mod-k': () => undefined } });
  const notes = await session.stories.footnotes();
  const holder = document.createElement('div');
  document.body.appendChild(holder);
  new StoryView(holder, session, notes[0]!, { label: 'Footnote 1' });
  readOnly = holder.querySelector('.ProseMirror')!;
});

afterEach(() => {
  resetHost();
  delete (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi;
});

/** VS Code's host page: the keys that reach its listener on `window`. */
function vscodeWindow(): { seen: string[]; stop: () => void } {
  const seen: string[] = [];
  const record = (e: KeyboardEvent): void => { seen.push(`${e.type} ${e.ctrlKey ? 'Ctrl+' : ''}${e.altKey ? 'Alt+' : ''}${e.shiftKey ? 'Shift+' : ''}${e.key}`); };
  window.addEventListener('keydown', record);
  window.addEventListener('keyup', record);
  return { seen, stop: () => { window.removeEventListener('keydown', record); window.removeEventListener('keyup', record); } };
}

function inVsCode(): void {
  initHost({ postMessage: () => undefined });
}

interface Press { key: string; ctrl?: boolean; alt?: boolean; shift?: boolean; keyCode?: number; type?: 'keydown' | 'keyup'; prevent?: boolean }

/** A key at `target`: a listener on the target prevents it first when `prevent`, as a handler of the page's would. */
function press(target: Element, p: Press): KeyboardEvent {
  const event = new KeyboardEvent(p.type ?? 'keydown', { key: p.key, ctrlKey: p.ctrl ?? false, altKey: p.alt ?? false, shiftKey: p.shift ?? false, keyCode: p.keyCode ?? 0, bubbles: true, cancelable: true });
  const prevent = (e: Event): void => e.preventDefault();
  if (p.prevent) target.addEventListener(event.type, prevent);
  target.dispatchEvent(event);
  if (p.prevent) target.removeEventListener(event.type, prevent);
  return event;
}

describe('in VS Code', () => {
  test('a key the page prevented is kept; Ctrl+S, which the page leaves, goes on', () => {
    inVsCode();
    const w = vscodeWindow();
    const button = document.body.appendChild(document.createElement('button'));
    try {
      press(button, { key: 'e', ctrl: true, shift: true, prevent: true });
      press(button, { key: 's', ctrl: true });
      expect(w.seen).toEqual(['keydown Ctrl+s']);
    } finally { w.stop(); button.remove(); }
  });

  test('in an editable surface, a key the page acted on is kept, and Ctrl+K, its shortcut; an unbound one goes on', () => {
    inVsCode();
    const w = vscodeWindow();
    try {
      // Bold acts at the caret, and the keymap prevents the key.
      const bold = press(body.view.dom, { key: 'b', ctrl: true, keyCode: 66 });
      expect(bold.defaultPrevented).toBe(true);
      // Ctrl+K, the application's shortcut on the surface.
      press(body.view.dom, { key: 'k', ctrl: true, keyCode: 75 });
      // Not bound: VS Code's.
      press(body.view.dom, { key: 's', ctrl: true, keyCode: 83 });
      press(body.view.dom, { key: 'p', ctrl: true, shift: true, keyCode: 80 });
      press(body.view.dom, { key: 'F1', keyCode: 112 });
      expect(w.seen).toEqual(['keydown Ctrl+s', 'keydown Ctrl+Shift+p', 'keydown F1']);
    } finally { w.stop(); }
  });

  test('a key the surface binds with Ctrl or Alt is kept though nothing prevented it (its command did nothing)', () => {
    /** An undispatched keydown at the body's element, as the gate would meet it with nothing prevented. */
    const at = (init: KeyboardEventInit): KeyboardEvent => {
      const event = new KeyboardEvent('keydown', { cancelable: true, ...init });
      Object.defineProperty(event, 'target', { value: body.view.dom });
      return event;
    };
    // Word's keys, the base keymap's and the application's shortcuts, as the surface installed them.
    for (const init of [{ key: 'b', ctrlKey: true }, { key: 'z', ctrlKey: true }, { key: '1', ctrlKey: true, altKey: true }, { key: 'Backspace', ctrlKey: true }, { key: 'k', ctrlKey: true }, { key: 'a', ctrlKey: true }]) {
      expect(pageKeepsKey(at(init)), JSON.stringify(init)).toBe(true);
    }
    // Not bound, or bound with no Ctrl, Cmd or Alt: the page leaves them.
    for (const init of [{ key: 's', ctrlKey: true }, { key: 'w', ctrlKey: true }, { key: 'P', ctrlKey: true, shiftKey: true }, { key: 'Tab' }, { key: 'Enter' }]) {
      expect(pageKeepsKey(at(init)), JSON.stringify(init)).toBe(false);
    }
    // The surface answers false while it is not editable.
    body.view.setProps({ editable: () => false });
    try { expect(pageKeepsKey(at({ key: 'b', ctrlKey: true }))).toBe(false); } finally { body.view.setProps({ editable: () => true }); }
  });

  test('in a read-only story view and outside any surface, Ctrl+B goes on: nothing there acts on it', () => {
    inVsCode();
    const w = vscodeWindow();
    const div = document.body.appendChild(document.createElement('div'));
    try {
      const inStory = press(readOnly, { key: 'b', ctrl: true, keyCode: 66 });
      expect(inStory.defaultPrevented).toBe(false);
      press(div, { key: 'b', ctrl: true, keyCode: 66 });
      expect(w.seen).toEqual(['keydown Ctrl+b', 'keydown Ctrl+b']);
    } finally { w.stop(); div.remove(); }
  });

  test('Alt+PageDown and Escape go on though the page acted on them, in a surface and in a dialog', () => {
    inVsCode();
    const w = vscodeWindow();
    const dialog = document.body.appendChild(document.createElement('div'));
    dialog.setAttribute('role', 'dialog');
    try {
      const before = stepped;
      const alt = press(body.view.dom, { key: 'PageDown', alt: true, keyCode: 34 });
      expect(alt.defaultPrevented).toBe(true);
      expect(stepped).toBe(before + 1);
      const escape = press(body.view.dom, { key: 'Escape', keyCode: 27 });
      press(dialog, { key: 'Escape', keyCode: 27, prevent: true });
      expect(w.seen).toEqual(['keydown Alt+PageDown', 'keydown Escape', 'keydown Escape']);
      // ProseMirror prevents every Escape (`captureKeyDown`): kept, VS Code could not hide its notification.
      expect(escape.defaultPrevented).toBe(true);
    } finally { w.stop(); dialog.remove(); }
  });

  test('in a field, its undo and redo are kept, so the field\'s own runs; its other keys go on', () => {
    inVsCode();
    const w = vscodeWindow();
    const input = document.body.appendChild(document.createElement('input'));
    const area = document.body.appendChild(document.createElement('textarea'));
    try {
      press(input, { key: 'z', ctrl: true, keyCode: 90 });
      press(input, { key: 'Z', ctrl: true, shift: true, keyCode: 90 });
      press(area, { key: 'y', ctrl: true, keyCode: 89 });
      press(input, { key: 'b', ctrl: true, keyCode: 66 });
      press(input, { key: 'k', ctrl: true, keyCode: 75 });
      expect(w.seen).toEqual(['keydown Ctrl+b', 'keydown Ctrl+k']);
    } finally { w.stop(); input.remove(); area.remove(); }
  });

  test('no keyup is ever kept', () => {
    inVsCode();
    const w = vscodeWindow();
    try {
      press(body.view.dom, { key: 'b', ctrl: true, keyCode: 66, type: 'keyup', prevent: true });
      press(body.view.dom, { key: 'PageDown', alt: true, keyCode: 34, type: 'keyup' });
      expect(w.seen).toEqual(['keyup Ctrl+b', 'keyup Alt+PageDown']);
    } finally { w.stop(); }
  });

  test('the gate goes with the host: reset, a prevented key reaches the window again', () => {
    inVsCode();
    resetHost();
    const w = vscodeWindow();
    try {
      press(body.view.dom, { key: 'b', ctrl: true, keyCode: 66 });
      expect(w.seen).toEqual(['keydown Ctrl+b']);
    } finally { w.stop(); }
  });
});

describe('in a browser', () => {
  test('no gate: every key reaches the window', () => {
    expect(initHost().kind).toBe('browser');
    const w = vscodeWindow();
    try {
      press(body.view.dom, { key: 'b', ctrl: true, keyCode: 66 });
      press(body.view.dom, { key: 'z', ctrl: true, keyCode: 90 });
      expect(w.seen).toEqual(['keydown Ctrl+b', 'keydown Ctrl+z']);
    } finally { w.stop(); }
  });
});

describe('the rule alone', () => {
  test('Alt with Shift is not Alt alone; a lone modifier key goes on', () => {
    const at = (): undefined => undefined;
    const kept = new KeyboardEvent('keydown', { key: 'x', altKey: true, shiftKey: true, cancelable: true });
    Object.defineProperty(kept, 'defaultPrevented', { value: true });
    Object.defineProperty(kept, 'target', { value: document.body });
    expect(pageKeepsKey(kept, at)).toBe(true);
    const alone = new KeyboardEvent('keydown', { key: 'Alt', altKey: true });
    Object.defineProperty(alone, 'target', { value: document.body });
    expect(pageKeepsKey(alone, at)).toBe(false);
  });
});

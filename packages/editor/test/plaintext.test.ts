// SPDX-License-Identifier: GPL-3.0-or-later
//
// Where a plain paste would go (ED-014 section 4.4, P34; section 9.15 item 1), decided before the key is taken and the
// clipboard read: nowhere for a button, a read-only or disabled field, a read-only story or a body with no active surface.
// jsdom has no `execCommand` or `ClipboardEvent`, so the pastes themselves are the browser suite's (`e2e/host-vscode.spec.ts`).
import { afterEach, describe, expect, test } from 'vitest';
import { isPlainPasteKey, plainPasteTarget } from '../src/app/plaintext.js';
import { registerSurfaceKeys, type SurfaceKeys } from '../src/view/surfacekeys.js';

const made: Element[] = [];
const add = <K extends keyof HTMLElementTagNameMap>(tag: K, setup: (e: HTMLElementTagNameMap[K]) => void = () => undefined): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  setup(e);
  document.body.appendChild(e);
  made.push(e);
  return e;
};
afterEach(() => { for (const e of made.splice(0)) e.remove(); });

const surface = (): SurfaceKeys & { pasted: string[] } => {
  const pasted: string[] = [];
  return { pasted, bindsKey: () => false, pasteText: (text) => { pasted.push(text); } };
};

describe('the plain paste\'s target', () => {
  test('a writable field takes it; a read-only or disabled one, a button and a read-only story take nothing', () => {
    expect(plainPasteTarget(add('input'), () => undefined)).toBeTypeOf('function');
    expect(plainPasteTarget(add('textarea'), () => undefined)).toBeTypeOf('function');
    expect(plainPasteTarget(add('textarea', (e) => { e.readOnly = true; }), () => undefined)).toBeUndefined();
    expect(plainPasteTarget(add('input', (e) => { e.disabled = true; }), () => undefined)).toBeUndefined();
    expect(plainPasteTarget(add('button'), () => surface())).toBeUndefined();
    const story = add('div', (e) => { e.className = 'ProseMirror'; e.contentEditable = 'false'; });
    expect(plainPasteTarget(story, () => surface())).toBeUndefined();
  });

  test('an editable surface takes it, from inside it; with the body focused, the active surface does, if there is one', () => {
    const view = add('div', (e) => { e.className = 'ProseMirror'; });
    const inner = view.appendChild(document.createElement('p'));
    const own = surface();
    const unregister = registerSurfaceKeys(view, own);
    try {
      expect(plainPasteTarget(inner, () => undefined)?.('in the surface')).toBe(true);
      expect(own.pasted).toEqual(['in the surface']);
      // Nothing pasted from an empty clipboard.
      expect(plainPasteTarget(inner, () => undefined)?.('')).toBe(false);
      const active = surface();
      expect(plainPasteTarget(document.body, () => active)?.('into the active one')).toBe(true);
      expect(active.pasted).toEqual(['into the active one']);
      expect(plainPasteTarget(document.body, () => null)).toBeUndefined();
    } finally { unregister(); }
  });

  test('the key: Ctrl+Shift+V and Ctrl+Alt+Shift+V (⌥⇧⌘V on a Mac), not Ctrl+V', () => {
    expect(isPlainPasteKey(new KeyboardEvent('keydown', { key: 'V', ctrlKey: true, shiftKey: true, keyCode: 86 }))).toBe(true);
    expect(isPlainPasteKey(new KeyboardEvent('keydown', { key: 'V', ctrlKey: true, altKey: true, shiftKey: true, keyCode: 86 }))).toBe(true);
    expect(isPlainPasteKey(new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, keyCode: 86 }))).toBe(false);
  });
});

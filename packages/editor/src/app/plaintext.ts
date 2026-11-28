// SPDX-License-Identifier: GPL-3.0-or-later
//
// Paste as plain text in VS Code (ED-014 section 4.4, P34). In a browser Chromium's own Ctrl+Shift+V hands the page the text
// alone; in VS Code's webview the host page cancels that key, and nothing replaces it. So there the page takes the key
// itself, reads the clipboard (VS Code grants a webview `clipboard-read`, measured in 1.147.0 and 1.105.0: section 9.13
// item 11) and pastes the text where its focus is. "Paste as Plain Text" from VS Code's menus comes as `pasteText`, the
// text read by the extension, and pastes the same way.
import { keyTest, surfaceKeysAt, type SurfaceKeys } from '../view/surfacekeys.js';

/** Ctrl+Shift+V, and ⌥⇧⌘V, Chrome's key on a Mac (⇧⌘V is the first name there). */
export const isPlainPasteKey = keyTest(['Mod-Shift-v', 'Mod-Alt-Shift-v']);

/**
 * Where a plain paste would go from `at` (the element focused when the paste was asked for), as a function that pastes the
 * text there and says whether it did; undefined where nothing can take it, so the key is left alone and the clipboard unread:
 * - a field, Monaco's included, takes it as its own typing (`insertText`), which the field can undo; a read-only or disabled
 *   one (a code pane before Edit) takes nothing;
 * - an editable surface takes it through ProseMirror's paste of the text and the paste rules (one web address a link, the
 *   Markdown offer), as in a browser;
 * - with no element focused (the page's `body`), `fallback`, the active surface, as VS Code's host page falls back to the
 *   element last focused for its own paste;
 * - anything else (a button, a read-only story) takes nothing.
 */
export function plainPasteTarget(at: Element | null, fallback: () => SurfaceKeys | null | undefined): ((text: string) => boolean) | undefined {
  if (at instanceof HTMLInputElement || at instanceof HTMLTextAreaElement) {
    if (at.readOnly || at.disabled) return undefined;
    return (text) => {
      if (text === '') return false;
      if (at.ownerDocument.activeElement !== at) at.focus();
      return at.ownerDocument.execCommand('insertText', false, text);
    };
  }
  const none = at === null || at === at.ownerDocument.body || at === at.ownerDocument.documentElement;
  const surface = none ? fallback() : surfaceKeysAt(at);
  if (!surface) return undefined;
  return (text) => {
    if (text === '') return false;
    surface.pasteText(text);
    return true;
  };
}

/** The text pasted as plain text where `at` is (`plainPasteTarget`); says whether it pasted. */
export function pastePlainText(text: string, at: Element | null, fallback: () => SurfaceKeys | null | undefined): boolean {
  return plainPasteTarget(at, fallback)?.(text) ?? false;
}

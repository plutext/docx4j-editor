// SPDX-License-Identifier: GPL-3.0-or-later
//
// The key gate (ED-014 section 4.4, P32): in VS Code, a key pressed in the page belongs to the page or to VS Code, never to
// both. VS Code's host page listens on the frame's `window` in the bubble phase and hands every keydown and keyup to the
// workbench, whose keybindings then run as if the key had been pressed there: Ctrl+B would bold and hide the side bar,
// Ctrl+Q clear the paragraph's formatting and quit (section 9.13). One listener on `document`, also in the bubble phase,
// stops a keydown the page keeps before the host page's listener sees it. It never calls `preventDefault()`, so the page's
// own handling stands, and it never touches a keyup, which VS Code's menu bar reads to tell a lone Alt (section 4.4).
//
// `initHost()` installs it when the host is VS Code's; the browser has no gate. A listener that acts on a key from anywhere
// in the page therefore listens in the capture phase and prevents the key when it acts: the gate never sees what a `window`
// listener does in the bubble phase, which runs after VS Code's.
import { surfaceKeysAt, keyTest, type SurfaceKeys } from '../view/surfacekeys.js';

/** A field's own undo and redo, which VS Code's host page would otherwise cancel. */
const fieldUndo = keyTest(['Mod-z', 'Mod-y', 'Mod-Shift-z']);

/** Where the gate finds an editable surface: by default the registry of the surfaces' `.ProseMirror` elements. */
export type SurfaceAt = (target: Element) => Pick<SurfaceKeys, 'bindsKey'> | undefined;

/**
 * Whether the page keeps this keydown from VS Code (P32):
 * - Escape goes on, wherever it is pressed: the page's uses of it close a dialog or the find bar, and VS Code may hide a
 *   notification as well.
 * - A key whose only modifier is Alt goes on (Alt+PageDown and Alt+PageUp, the next and the previous change): kept, it would
 *   leave VS Code's menu bar seeing Alt pressed alone, and the menu bar would take the focus (section 9.13 item 3).
 * - A key the page acted on (prevented) is kept.
 * - In a field, its undo and redo are kept, so the field's own runs.
 * - In an editable surface, a key the surface binds with Ctrl, Cmd or Alt is kept, whether or not its command acted: Ctrl+B
 *   in a locked control stays the page's and does nothing, as in Word.
 * - Every other key goes on, as in any webview: Ctrl+S, Ctrl+P, Ctrl+Shift+P, F1, Ctrl+W, and copy, cut and paste, which
 *   VS Code hands back through `execCommand`.
 */
export function pageKeepsKey(event: KeyboardEvent, surfaceAt: SurfaceAt = surfaceKeysAt): boolean {
  if (event.key === 'Escape') return false;
  if (event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) return false;
  if (event.defaultPrevented) return true;
  const target = event.target instanceof Element ? event.target : null;
  if (!target) return false;
  if ((target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) && fieldUndo(event)) return true;
  if (!(event.ctrlKey || event.metaKey || event.altKey)) return false;
  return surfaceAt(target)?.bindsKey(event) ?? false;
}

/** Installs the gate on the page's `document`; returns its removal. */
export function installKeyGate(doc: Document = document, surfaceAt: SurfaceAt = surfaceKeysAt): () => void {
  const onKey = (event: KeyboardEvent): void => {
    if (pageKeepsKey(event, surfaceAt)) event.stopPropagation();
  };
  doc.addEventListener('keydown', onKey);
  return () => doc.removeEventListener('keydown', onKey);
}

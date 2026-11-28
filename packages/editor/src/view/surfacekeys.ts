// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the page's key gate and its plain paste ask of an editing surface (ED-014 section 4.4), found by the surface's
// `.ProseMirror` element, and the key test the surfaces' keymaps answer through. Light (prosemirror-keymap's handler alone),
// so the gate (`app/keygate.ts`) loads without the view. Each editable surface registers itself when it mounts
// (`SurfaceBase.mount`); a read-only story view never does.
import type { Command } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { keydownHandler } from 'prosemirror-keymap';

export interface SurfaceKeys {
  /** Whether the surface binds this key in one of its keymaps, whether or not the key's command would act now. */
  bindsKey(event: KeyboardEvent): boolean;
  /** The text pasted as plain text where the surface's selection is, through the surface's paste rules. */
  pasteText(text: string): void;
}

const surfaces = new WeakMap<Element, SurfaceKeys>();

/** Registers a surface under its view's element; returns the unregister. */
export function registerSurfaceKeys(element: Element, surface: SurfaceKeys): () => void {
  surfaces.set(element, surface);
  return () => { if (surfaces.get(element) === surface) surfaces.delete(element); };
}

/** The editable surface an element is in: the nearest `.ProseMirror` around it, if one is registered. */
export function surfaceKeysAt(target: Element | null): SurfaceKeys | undefined {
  const element = target?.closest('.ProseMirror');
  return element ? surfaces.get(element) : undefined;
}

/** A key test over binding names: whether a keydown matches one, as prosemirror-keymap matches it. */
export type KeyTest = (event: KeyboardEvent) => boolean;

/**
 * The view prosemirror-keymap's handler is given here. The handler matches over the event and the table alone, and passes
 * the view only to the binding it finds, which only answers, so no view is read (prosemirror-keymap 1.2: `view.state` and
 * `view.dispatch` as the binding's arguments). The one stand-in, so that a release that reads the view fails here, once,
 * and in `keygate.test.ts`.
 */
const VIEWLESS = Object.freeze({}) as EditorView;

/**
 * Whether a keydown matches one of these names, as a keymap of them would (Shift, `Mod`, a shifted character, the key code's
 * fallback, AltGr on Windows): prosemirror-keymap's own handler over bindings that only answer, so it never runs a command.
 */
export function keyTest(names: Iterable<string>): KeyTest {
  const answer: Command = () => true;
  const handler = keydownHandler(Object.fromEntries([...names].map((name) => [name, answer])));
  return (event) => handler(VIEWLESS, event);
}

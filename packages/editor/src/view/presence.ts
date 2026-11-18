// SPDX-License-Identifier: GPL-3.0-or-later
//
// Presence (ED-004 section 5): the people in a shared room, their carets and selections. The
// room's awareness (the provider's) is registered per session; every editing surface, the body
// and each editable note or header, carries `presencePlugin` from its creation, which follows the
// registration: it draws the peers' carets while the session is in a room and writes this peer's
// own, and draws nothing otherwise. A caret is one widget and a selection one inline decoration
// per peer, a handful however long the document (ED-002 section 10.5 item 10).
//
// Not y-prosemirror's `yCursorPlugin` installed on joining: adding a plugin to a live view changes
// its plugin list, which makes ProseMirror destroy and remake every plugin view, and
// `yUndoPlugin`'s destroy destroys the session's undo manager (upstream item 4), so undo stopped
// working in a room (ED-004 section 10.8 item 1). The plugin below does what `yCursorPlugin` does
// with the pieces y-prosemirror exports (`createDecorations`, the relative-position conversions),
// and its redraws go through a timer of its own that a destroyed view never receives.
//
// UPSTREAM-7: y-prosemirror 1.3.7 draws the peers' carets from `aw.cursor` alone
// (`createDecorations`), so a field per story (section 5's `cursor:<story key>`) could not be
// drawn. Every surface uses the one `cursor` field. That is enough: a caret is a Yjs relative
// position, which resolves only inside its own fragment, so each surface draws the carets in its
// own story and no other, and the story a peer is in is read from the fragment the position
// resolves to (`whereIs`).
import type { Awareness } from 'y-protocols/awareness';
import * as Y from 'yjs';
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { DecorationSet } from 'prosemirror-view';
import { ySyncPluginKey, createDecorations, defaultAwarenessStateFilter, absolutePositionToRelativePosition, relativePositionToAbsolutePosition } from 'y-prosemirror';
import type { EditStore, EditorDocument } from '@docx4j/editor-model';
import { markupKey, clampCaret } from './markup.js';

const registry = new WeakMap<EditorDocument, Awareness>();
const listeners = new WeakMap<EditorDocument, Set<() => void>>();

/** The room's awareness for a session's surfaces; null when the session leaves the room. */
export function setPresence(session: EditorDocument, awareness: Awareness | null): void {
  if (awareness) registry.set(session, awareness); else registry.delete(session);
  for (const fn of [...(listeners.get(session) ?? [])]) fn();
}

export function presenceOf(session: EditorDocument): Awareness | undefined {
  return registry.get(session);
}

/** Told when the session's awareness is registered or dropped. */
export function onPresence(session: EditorDocument, fn: () => void): () => void {
  let set = listeners.get(session);
  if (!set) { set = new Set(); listeners.set(session, set); }
  set.add(fn);
  return () => set!.delete(fn);
}

interface PresenceState {
  awareness: Awareness | null;
  set: DecorationSet;
}

export const presenceKey = new PluginKey<PresenceState>('docx4j-presence');

interface SyncState {
  type: Y.XmlFragment;
  doc: Y.Doc;
  binding?: { mapping: Map<unknown, unknown> } | null;
  isChangeOrigin?: boolean;
}

/**
 * The peers' states with each collapsed caret where this view draws a caret (ED-005 section 6.4):
 * one inside content the view hides (As accepted, Original) is moved to the run's drawn edge, as
 * a local caret is, since a widget there is inside hidden DOM and draws nothing (the spike,
 * section 12.9). `createDecorations` reads only `getStates()`.
 */
function clampedStates(state: EditorState, awareness: Awareness, sync: SyncState): Pick<Awareness, 'getStates'> {
  if ((markupKey.getState(state)?.view ?? 'all') === 'all') return awareness;
  const mapping = sync.binding!.mapping as never;
  const out = new Map<number, Record<string, unknown>>();
  awareness.getStates().forEach((aw, id) => {
    const cursor = (aw as { cursor?: { anchor: unknown; head: unknown } | null }).cursor;
    const at = (rel: unknown): number | null => relativePositionToAbsolutePosition(sync.doc, sync.type, Y.createRelativePositionFromJSON(rel), mapping);
    const head = cursor ? at(cursor.head) : null;
    if (!cursor || head === null || at(cursor.anchor) !== head) { out.set(id, aw); return; }
    const drawn = clampCaret(state, head);
    if (drawn === head) { out.set(id, aw); return; }
    const rel = absolutePositionToRelativePosition(drawn, sync.type, mapping);
    out.set(id, { ...aw, cursor: { anchor: rel, head: rel } });
  });
  return { getStates: () => out };
}

function redraw(state: EditorState, awareness: Awareness | null): DecorationSet {
  const sync = ySyncPluginKey.getState(state) as SyncState | undefined;
  if (!awareness || !sync?.binding) return DecorationSet.empty;
  return createDecorations(state, clampedStates(state, awareness, sync) as never, defaultAwarenessStateFilter, caretBuilder as never, selectionBuilder as never) as DecorationSet;
}

/**
 * The peers' carets on a surface, and this peer's caret written to the room, while the session
 * is in a room. Carried from the surface's creation, so joining and leaving never change the
 * plugin list (see the header).
 */
export function presencePlugin(session: EditorDocument): Plugin<PresenceState> {
  return new Plugin<PresenceState>({
    key: presenceKey,
    state: {
      init: () => ({ awareness: null, set: DecorationSet.empty }),
      apply(tr, previous, _old, next) {
        const meta = tr.getMeta(presenceKey) as { awareness: Awareness | null } | undefined;
        const awareness = meta ? meta.awareness : previous.awareness;
        const sync = ySyncPluginKey.getState(next) as SyncState | undefined;
        // A change of the markup view moves the carets it hides (ED-005 section 6.4).
        if (meta || sync?.isChangeOrigin || tr.getMeta(markupKey)) return { awareness, set: redraw(next, awareness) };
        return { awareness, set: previous.set.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => presenceKey.getState(state)?.set,
    },
    view(view) {
      let destroyed = false;
      let scheduled = false;
      let attached: Awareness | null = null;
      // A redraw on the next tick, as y-prosemirror's own does, so an awareness change inside a
      // view update never dispatches from within it; never to a view destroyed meanwhile.
      const schedule = (): void => {
        if (scheduled) return;
        scheduled = true;
        setTimeout(() => {
          scheduled = false;
          if (!destroyed) view.dispatch(view.state.tr.setMeta(presenceKey, { awareness: attached }));
        }, 0);
      };
      const attach = (): void => {
        const awareness = presenceOf(session) ?? null;
        if (awareness === attached) return;
        attached?.off('change', schedule);
        attached = awareness;
        attached?.on('change', schedule);
        schedule();
      };
      /** This peer's caret into the room while the surface has focus, taken back when it loses it (as `yCursorPlugin` does). */
      const writeCaret = (): void => {
        const sync = ySyncPluginKey.getState(view.state) as SyncState | undefined;
        if (!attached || !sync?.binding || destroyed) return;
        const mapping = sync.binding.mapping as never;
        const current = (attached.getLocalState() ?? {}) as { cursor?: { anchor: unknown; head: unknown } | null };
        if (view.hasFocus()) {
          const { anchor: a, head: h } = view.state.selection;
          const anchor = absolutePositionToRelativePosition(a, sync.type, mapping);
          const head = absolutePositionToRelativePosition(h, sync.type, mapping);
          if (!current.cursor
            || !Y.compareRelativePositions(Y.createRelativePositionFromJSON(current.cursor.anchor), anchor)
            || !Y.compareRelativePositions(Y.createRelativePositionFromJSON(current.cursor.head), head)) {
            attached.setLocalStateField('cursor', { anchor, head });
          }
        } else if (current.cursor
          && relativePositionToAbsolutePosition(sync.doc, sync.type, Y.createRelativePositionFromJSON(current.cursor.anchor), mapping) !== null) {
          // The caret this surface wrote, taken back: another surface writes its own on focus.
          attached.setLocalStateField('cursor', null);
        }
      };
      const offPresence = onPresence(session, attach);
      attach();
      view.dom.addEventListener('focusin', writeCaret);
      view.dom.addEventListener('focusout', writeCaret);
      return {
        update: writeCaret,
        destroy: () => {
          destroyed = true;
          offPresence();
          attached?.off('change', schedule);
          view.dom.removeEventListener('focusin', writeCaret);
          view.dom.removeEventListener('focusout', writeCaret);
        },
      };
    },
  });
}

/** What a peer tells the others about itself, in the awareness's `user` field. */
export interface PeerUser {
  name: string;
  /** `#rrggbb`: y-prosemirror draws only that form. */
  color: string;
}

/** A peer's caret: a bar in its colour, its name above it on hover; `data-peer` names the client for the people list's "go to". */
export function caretBuilder(user: PeerUser, clientId: number): HTMLElement {
  const caret = document.createElement('span');
  caret.className = 'peer-caret';
  caret.dataset['peer'] = String(clientId);
  caret.style.borderColor = user.color;
  const label = document.createElement('span');
  label.className = 'peer-caret-name';
  label.style.backgroundColor = user.color;
  label.textContent = user.name;
  caret.append('⁠', label, '⁠');
  return caret;
}

/** A peer's selection: a tint of its colour. */
export function selectionBuilder(user: PeerUser): { class: string; style: string } {
  return { class: 'peer-selection', style: `background-color: ${user.color}33` };
}

/** A colour for a name, the comments margin's hue (`authorHue`) as `#rrggbb`. */
export function peerColor(hue: number): string {
  // hsl(h 65% 42%), as `authorColor` draws an author.
  const s = 0.65;
  const l = 0.42;
  const k = (n: number): number => (n + hue / 30) % 12;
  const a = s * Math.min(l, 1 - l);
  const f = (n: number): number => l - a * Math.max(-1, Math.min(k(n) - 3, 9 - k(n), 1));
  const hex = (x: number): string => Math.round(x * 255).toString(16).padStart(2, '0');
  return `#${hex(f(0))}${hex(f(8))}${hex(f(4))}`;
}

/**
 * Which story a caret is in: `body`, `footnote:<id>`, `endnote:<id>`, `header:<part>` or
 * `footer:<part>`; undefined when it resolves nowhere (a peer that has left the text, or a
 * position in a fragment this store lacks).
 */
export function whereIs(store: EditStore, cursor: unknown): string | undefined {
  const head = (cursor as { head?: unknown } | null | undefined)?.head;
  if (!head) return undefined;
  let absolute: Y.AbsolutePosition | null;
  try {
    absolute = Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(head), store.ydoc);
  } catch {
    return undefined;
  }
  if (!absolute) return undefined;
  // Up from the type the position is in to the story's fragment: the body (a root type, no
  // parent) or a value of one of the story maps. Not the first `XmlFragment` on the way:
  // `Y.XmlElement` extends it, so a paragraph would stop the walk.
  let type: Y.AbstractType<unknown> | null = absolute.type;
  while (type && type !== store.body && !(type.parent instanceof Y.Map)) type = type.parent as Y.AbstractType<unknown> | null;
  if (!type) return undefined;
  if (type === store.body) return 'body';
  const inMap = (map: Y.Map<Y.XmlFragment>, kind: (key: string) => string): string | undefined => {
    for (const [key, value] of map.entries()) if (value === type) return kind(key);
    return undefined;
  };
  return inMap(store.footnotes, (k) => `footnote:${k}`)
    ?? inMap(store.endnotes, (k) => `endnote:${k}`)
    ?? inMap(store.headers, (k) => `${/footer/i.test(k) ? 'footer' : 'header'}:${k}`);
}

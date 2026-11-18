// SPDX-License-Identifier: GPL-3.0-or-later
//
// What first marked a document as having unsaved changes (ED-011 section 9.9; Jason, 2026-11-09): a tab said a
// document had unsaved changes with nothing edited there, and by the time it was looked at its session had been
// closed and deleted, so nothing showed what had counted as an edit. The first store transaction the dirty state
// counts is described here and kept twice: in the session's record, and in a short log in `localStorage` that
// outlives the session, since a Close, a Reload or a Discard deletes the record. Nothing of the document's text is
// kept, only where the change was and what led to it.
import * as Y from 'yjs';

/** The first edit since the open or the last save, described for a person reading it afterwards. */
export interface EditCause {
  /** When, as an ISO time. */
  at: string;
  /** The transaction's origin: y-prosemirror's key (`y-sync$`) for a view's edit, `undo`, `local` for a command run headless, a peer's provider. */
  origin: string;
  /** Whether Yjs counts the transaction as this tab's own. */
  local: boolean;
  /** The store's top-level entries it changed: `body`, `package`, `parts`, a story's map. */
  changed: string[];
  /** The body's top-level blocks it changed, by index, the first few; absent when it changed none it can name. */
  blocks?: number[];
  /** The last input the page saw before it, if any: what it was and how long before. */
  input?: InputSeen;
  /** The element that had the focus: a test id, else its tag and first class. */
  focus?: string;
  /** Whether the page was visible. */
  visible: boolean;
  /** The calls that made the transaction, its first frames, the page's address left out. */
  stack?: string[];
}

/** An input event, as the cause names it: a key by its name (a character is `character`, never which one). */
export interface InputSeen {
  kind: string;
  key?: string;
  msBefore: number;
  target?: string;
}

const LOG_KEY = 'docx4j-editor.editCauses';
const LOG_SIZE = 20;
const BLOCKS = 5;
const FRAMES = 30;

/** An entry of the log: the cause with the session it marked. */
export interface EditCauseEntry extends EditCause {
  session: string;
  name: string;
}

/** The origin as a word: a string as it is, a plugin key by its key, the undo manager as `undo`, anything else by its class. */
export function describeOrigin(origin: unknown): string {
  if (origin === null || origin === undefined) return 'none';
  if (typeof origin === 'string') return origin;
  if (origin instanceof Y.UndoManager) return 'undo';
  const key = (origin as { key?: unknown }).key;
  if (typeof key === 'string') return key;
  const name = (origin as { constructor?: { name?: unknown } }).constructor?.name;
  return typeof name === 'string' && name !== '' ? name : typeof origin;
}

/** The root a type hangs from, and its child of that root on the way up. Types are compared by identity alone, so held as `unknown`. */
function rootAndChild(type: unknown): { root: unknown; child?: unknown } {
  let child: unknown;
  let t = type;
  for (;;) {
    const parent = (t as { _item?: { parent?: unknown } | null })._item?.parent;
    if (!(parent instanceof Y.AbstractType)) return { root: t, ...(child !== undefined ? { child } : {}) };
    child = t;
    t = parent;
  }
}

/** The changes a transaction made to the store, described: which entries, and which of the body's blocks. */
export function describeChanges(tr: Y.Transaction, body: Y.XmlFragment): Pick<EditCause, 'origin' | 'local' | 'changed' | 'blocks'> {
  const names = new Map<unknown, string>();
  for (const [name, type] of tr.doc.share) names.set(type, name);
  const changed = new Set<string>();
  const children = new Set<unknown>();
  for (const type of tr.changed.keys()) {
    const { root, child } = rootAndChild(type);
    changed.add(names.get(root) ?? 'unnamed');
    // A block put in or taken out changes the body itself, and is named by no child: the list then says less.
    if (root === body && child !== undefined) children.add(child);
  }
  const blocks: number[] = [];
  if (children.size > 0) {
    // Once per cause, never per keystroke: the first edit since the last save.
    body.toArray().forEach((block, i) => { if (blocks.length < BLOCKS && children.has(block)) blocks.push(i); });
  }
  return { origin: describeOrigin(tr.origin), local: tr.local, changed: [...changed].sort(), ...(blocks.length > 0 ? { blocks } : {}) };
}

/**
 * The stack where the call is made, deep enough to reach past Yjs and ProseMirror to what dispatched: V8 keeps ten
 * frames unless told otherwise, and the dirty state's listener sits about fifteen below the edit's own code.
 */
export function deepStack(): string | undefined {
  const holder = Error as { stackTraceLimit?: number };
  const limit = holder.stackTraceLimit;
  holder.stackTraceLimit = 40;
  try { return new Error().stack; } finally { holder.stackTraceLimit = limit; }
}

/** The call stack's frames, with the page's address taken out of each. */
export function stackFrames(stack: string | undefined, origin: string = typeof location !== 'undefined' ? location.origin : ''): string[] | undefined {
  if (!stack) return undefined;
  const lines = stack.split('\n').map((l) => l.trim()).filter((l) => l !== '' && l !== 'Error' && !l.startsWith('Error:'));
  const frames = lines.map((l) => (origin ? l.split(origin).join('') : l));
  return frames.length > 0 ? frames.slice(0, FRAMES) : undefined;
}

/** An element as the cause names it: its test id, else its tag and first class. */
export function describeElement(el: Element | null | undefined): string | undefined {
  if (!el) return undefined;
  const id = el.closest('[data-testid]')?.getAttribute('data-testid');
  const inEditor = el.closest('.ProseMirror')?.getAttribute('aria-label');
  if (inEditor) return `editor:${inEditor}`;
  if (id) return id;
  const cls = typeof el.className === 'string' ? el.className.split(/\s+/)[0] : '';
  return `${el.tagName.toLowerCase()}${cls ? '.' + cls : ''}`;
}

/** A key as the cause names it: a named key with its modifiers, or `character`, so no typed text is kept. */
export function describeKey(e: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey'>): string {
  const named = e.key.length > 1 ? e.key : (e.ctrlKey || e.metaKey || e.altKey ? e.key.toUpperCase() : 'character');
  const mods = [e.ctrlKey ? 'Ctrl' : '', e.metaKey ? 'Meta' : '', e.altKey ? 'Alt' : '', e.shiftKey && named !== 'character' ? 'Shift' : ''].filter((m) => m !== '');
  return [...mods, named].join('+');
}

interface Seen { kind: string; key?: string; time: number; target?: string }

/**
 * The page's last input, followed in the capture phase from `start` until the function it returns is called: a key,
 * a pointer, a paste, a drop, a cut, the end of a composition. Its `seen` reads it as a cause's `input`.
 */
export function followInput(win: Window = window): { seen: (now?: number) => InputSeen | undefined; stop: () => void } {
  let last: Seen | undefined;
  const note = (e: Event): void => {
    const target = describeElement(e.target instanceof Element ? e.target : null);
    const key = e instanceof KeyboardEvent ? describeKey(e) : undefined;
    last = { kind: e.type, time: performance.now(), ...(key ? { key } : {}), ...(target ? { target } : {}) };
  };
  const kinds = ['keydown', 'pointerdown', 'paste', 'drop', 'cut', 'compositionend'] as const;
  for (const kind of kinds) win.addEventListener(kind, note, { capture: true, passive: true });
  return {
    seen: (now = performance.now()) => (last ? { kind: last.kind, msBefore: Math.round(now - last.time), ...(last.key ? { key: last.key } : {}), ...(last.target ? { target: last.target } : {}) } : undefined),
    stop: () => { for (const kind of kinds) win.removeEventListener(kind, note, { capture: true }); },
  };
}

/** The whole cause of a transaction the dirty state counted, read at that moment. */
export function editCauseOf(tr: Y.Transaction, body: Y.XmlFragment, input: InputSeen | undefined, stack: string | undefined): EditCause {
  const focus = typeof document !== 'undefined' ? describeElement(document.activeElement) : undefined;
  const frames = stackFrames(stack);
  return {
    at: new Date().toISOString(),
    ...describeChanges(tr, body),
    ...(input ? { input } : {}),
    ...(focus ? { focus } : {}),
    visible: typeof document === 'undefined' || document.visibilityState === 'visible',
    ...(frames ? { stack: frames } : {}),
  };
}

/** The log of the last causes, oldest first; empty where storage is refused. */
export function editCauseLog(): EditCauseEntry[] {
  try {
    const raw = localStorage.getItem(LOG_KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(parsed) ? (parsed as EditCauseEntry[]) : [];
  } catch { return []; }
}

/** A cause added to the log, the oldest let go past twenty. Storage refused, nothing is kept and nothing fails. */
export function logEditCause(entry: EditCauseEntry): void {
  try {
    localStorage.setItem(LOG_KEY, JSON.stringify([...editCauseLog(), entry].slice(-LOG_SIZE)));
  } catch { /* a private window, or storage full: the session record still has it */ }
}

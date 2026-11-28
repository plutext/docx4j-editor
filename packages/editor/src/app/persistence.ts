// SPDX-License-Identifier: GPL-3.0-or-later
//
// Autosave and recovery (ED-001 section 11 item 1; ED-002 section 5.4): the Yjs update log
// through y-indexeddb, and the original bytes and name in an IndexedDB store of our own, keyed
// by a session id, because the Yjs state alone cannot export. On the next visit an unsaved
// session is offered for recovery: the bytes are opened afresh and the persisted store adopted.
import type { KeptPlace } from '../view/place.js';
import { IndexeddbPersistence } from 'y-indexeddb';
import * as Y from 'yjs';
import { newEditStore, type EditStore } from '@docx4j/editor-model';
import type { RoomLink } from './room.js';
import type { EditCause } from './editcause.js';

const DB = 'docx4j-editor';
const STORE = 'sessions';
/** The hand-over records of ED-011 section 3.6: the disk version for a new tab, keyed by a fresh id, taken once. */
const HANDOVERS = 'handovers';
const CURRENT = 'docx4j-editor.session';

export interface SessionRecord {
  id: string;
  name: string;
  bytes: Uint8Array;
  /** Unsaved changes since the last successful save or the open. */
  dirty: boolean;
  updated: number;
  /** The room the session is in (ED-004 section 4.5): recovery rejoins it. */
  room?: RoomLink;
  /** Whether this peer joined the room (its bytes are the room's base file) rather than shared it. */
  joined?: boolean;
  /** While dirty: what first counted as an edit since the open or the last save (ED-011 section 9.9). */
  firstEdit?: EditCause;
}

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 2);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
      if (!db.objectStoreNames.contains(HANDOVERS)) db.createObjectStore(HANDOVERS, { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function request<T>(r: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => { r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
}

export async function putSession(record: SessionRecord): Promise<void> {
  const db = await openDb();
  await request(db.transaction(STORE, 'readwrite').objectStore(STORE).put(record));
  db.close();
}

export async function getSession(id: string): Promise<SessionRecord | undefined> {
  const db = await openDb();
  const out = await request<SessionRecord | undefined>(db.transaction(STORE, 'readonly').objectStore(STORE).get(id));
  db.close();
  return out;
}

export async function deleteSession(id: string): Promise<void> {
  // The pointer goes first, synchronously: a reload during the deletes below must not find the session.
  try { if (localStorage.getItem(CURRENT) === id) localStorage.removeItem(CURRENT); } catch { /* ignore */ }
  const db = await openDb();
  await request(db.transaction(STORE, 'readwrite').objectStore(STORE).delete(id));
  db.close();
  await new Promise<void>((resolve) => { const r = indexedDB.deleteDatabase(`${DB}:${id}`); r.onsuccess = () => resolve(); r.onerror = () => resolve(); r.onblocked = () => resolve(); });
}

export function currentSessionId(): string | null {
  try { return localStorage.getItem(CURRENT); } catch { return null; }
}

export function setCurrentSessionId(id: string | null): void {
  try { if (id === null) localStorage.removeItem(CURRENT); else localStorage.setItem(CURRENT, id); } catch { /* ignore */ }
}

export function newSessionId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/** A store persisted under the session id; resolves once IndexedDB has been read into it. */
export async function persistedStore(id: string): Promise<{ store: EditStore; persistence: IndexeddbPersistence }> {
  const ydoc = new Y.Doc();
  const persistence = new IndexeddbPersistence(`${DB}:${id}`, ydoc);
  await persistence.whenSynced;
  return { store: newEditStore(ydoc), persistence };
}

/** What keeps a store, as the application holds it: y-indexeddb's persistence, or nothing. */
export interface StoreKeeper {
  destroy(): void;
}

/**
 * A store kept in memory alone, for a host whose own backups hold the session (VS Code's hot exit, ED-014 P8 and P20);
 * with `update`, a backup's state applied first, as y-indexeddb applies a kept one.
 */
export function memoryStore(update?: Uint8Array): { store: EditStore; persistence: StoreKeeper } {
  const ydoc = new Y.Doc();
  if (update && update.byteLength > 0) Y.applyUpdate(ydoc, update);
  return { store: newEditStore(ydoc), persistence: { destroy: () => undefined } };
}

/**
 * How long a check waits for a tab that is going away. Chromium and Firefox had let go before the next page's first
 * script in all 16 reloads measured, but nothing promises that order, and an offer missed on a reload leaves a person
 * at an empty page with their work out of sight. The wait is unseen: a tab that stays open means no offer either way.
 */
const LEAVING = 1000;

const sessionLock = (id: string): string => `${DB}.session.${id}`;

/**
 * Mark a session as open in this tab, until the function returned is called (ED-011 section 3.6). A lock, since the
 * browser lets it go when the tab closes or crashes, where a flag in storage would still say "open" after the very crash
 * recovery is for. Shared, as two tabs of one browser may join one room under one session id. Without Web Locks (a page
 * on an address that is not secure) nothing is held, and every session reads as free, as it did before.
 */
export function holdSession(id: string): () => void {
  if (!navigator.locks) return () => undefined;
  let release = (): void => undefined;
  const held = new Promise<void>((resolve) => { release = resolve; });
  navigator.locks.request(sessionLock(id), { mode: 'shared' }, () => held).catch(() => undefined);
  return release;
}

/** Whether a tab has this session open. Asked for rather than queried, so that a tab going away has its moment to go. */
export async function sessionInUse(id: string): Promise<boolean> {
  if (!navigator.locks) return false;
  try {
    await navigator.locks.request(sessionLock(id), { signal: AbortSignal.timeout(LEAVING) }, () => undefined);
    return false;
  } catch (e) {
    return (e as Error).name === 'TimeoutError';
  }
}

/** The unsaved session to offer for recovery, if any: a session another tab has open is not lost. Read after the check, since a tab that was closing may have deleted it. */
export async function recoverable(): Promise<SessionRecord | undefined> {
  const id = currentSessionId();
  if (!id) return undefined;
  try {
    if (await sessionInUse(id)) return undefined;
    const record = await getSession(id);
    return record?.dirty ? record : undefined;
  } catch {
    return undefined;
  }
}

/**
 * The disk version handed to a new tab (ED-011 section 3.6): the file's name, bytes and modification time as the
 * handing tab read them, and the handle where it can be stored (a `FileSystemFileHandle` is structured-cloneable
 * in Chromium; the new tab asks its permission from a click, P5).
 */
export interface HandoverRecord {
  id: string;
  name: string;
  bytes: Uint8Array;
  lastModified: number;
  handle?: FileSystemFileHandle;
  /** The bytes are a Markdown source's: the new tab converts them and opens the result (ED-011 section 3.8). */
  markdown?: true;
  /** The handing tab's caret and view, found again in the disk version (ED-011 section 3.5). */
  place?: KeptPlace;
}

/** Write a hand-over record; a handle the store cannot clone is left out and the new tab opens the bytes alone. */
export async function putHandover(record: HandoverRecord): Promise<void> {
  const db = await openDb();
  try {
    const store = db.transaction(HANDOVERS, 'readwrite').objectStore(HANDOVERS);
    try { await request(store.put(record)); }
    catch (e) {
      if ((e as Error).name !== 'DataCloneError' || !record.handle) throw e;
      const { handle: _handle, ...bare } = record;
      await request(db.transaction(HANDOVERS, 'readwrite').objectStore(HANDOVERS).put(bare));
    }
  } finally { db.close(); }
}

/** Read a hand-over record and delete it: a tab reloaded finds nothing and opens nothing. */
export async function takeHandover(id: string): Promise<HandoverRecord | undefined> {
  const db = await openDb();
  try {
    const tx = db.transaction(HANDOVERS, 'readwrite');
    const store = tx.objectStore(HANDOVERS);
    const out = await request<HandoverRecord | undefined>(store.get(id));
    if (out) await request(store.delete(id));
    return out;
  } finally { db.close(); }
}

/** The id in a `#disk=<id>` address, the new tab's way in (ED-011 section 3.6). */
export function parseHandoverHash(hash: string): string | undefined {
  const m = /^#disk=([A-Za-z0-9-]+)$/.exec(hash);
  return m ? m[1] : undefined;
}

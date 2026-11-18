// SPDX-License-Identifier: GPL-3.0-or-later
//
// The room repair (ED-005 section 12.7 item 5 (b), decided by Jason 2026-09-28; section 12.15).
// UPSTREAM-10: Yjs keeps formatting as markers at positions (docs/developer/upstream-limitations.md item 10).
// Yjs keeps formatting as markers at positions, and text typed inside a range whose attributes
// equal its own gets no marker of its own (docs/developer/upstream-limitations.md item 10). So an accept or
// a reject at the moment a peer types inside the change reaches the typing: an accept takes it in
// unseen, a reject leaves it plain. And typing with a marker of its own ends with a marker that
// restores what its writer saw, which can put the resolved mark back on text after it.
//
// What a resolution writes. The four commands, in a room, carry a record on their transaction
// (`addResolutionWrite`), which the view and the headless runner write first in the same Yjs
// transaction as the resolution: the state vector the resolver saw, and each resolved text piece's
// range with its mark's key and attributes. The range is two relative positions, its start after
// the character before it and its end before the character after it, so that text typed at either
// edge while it was resolved lies inside.
//
// What every peer repairs, after any remote update, the same way from the same state:
//
// - text the resolution saw (its item covered by the state vector) inside a range and still
//   carrying the resolved mark has the resolution applied again (the mark taken off), which is what
//   a concurrent typist's restoring marker undoes;
// - text it did not see, typed concurrently inside a range, carrying no marker of its own for the
//   key and none of the key now, carried the change's own attributes when it was typed (Yjs writes a
//   marker for any other value), so the change's mark goes back on it, with the change's own
//   author and date.
//
// Concurrent is bounded by acknowledgements: each peer records, in the room, its own clock when a
// resolution first reaches it (`resolutionAcks`), and text it typed after that saw the resolution
// and is never repaired. Until a peer's acknowledgement has arrived, all its text the resolution
// did not see is concurrent, which holds, since Yjs delivers a client's edits in its order and its
// later text comes after its acknowledgement.
//
// Termination: a repair writes only where one of the two conditions holds, and after it neither
// holds for those characters; it writes formatting, which is no text for either condition to hold
// of; it runs on remote updates alone, its own writes being local; two peers repairing at once
// write equal values over the same characters, which Yjs merges to those values. So a second run
// over any state a repair left writes nothing.
/* eslint-disable @typescript-eslint/no-explicit-any -- Yjs's types are generic over their events, and a walk over them takes any. */
import * as Y from 'yjs';
import * as sha256 from 'lib0/hash/sha256';
import * as buf from 'lib0/buffer';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Mark, Schema } from 'prosemirror-model';
import { absolutePositionToRelativePosition, initProseMirrorDoc } from 'y-prosemirror';
import type { EditStore } from './yjs.mjs';
import type { TrackedChangeInfo } from './review.mjs';
import { plainEqual } from './json.mjs';

type RevisionKey = 'insertion' | 'deletion' | 'formatChange';

/** One resolved text piece: its range in the store, its mark and the attribute key the store holds it by, and its attributes. */
export interface ResolvedRange {
  from: unknown;
  to: unknown;
  mark: RevisionKey;
  /** The Yjs attribute's key: the mark's name, or for a deletion `deletion--<hash>` (`yAttributeKey`). */
  key: string;
  attrs: Record<string, unknown>;
}

/**
 * The key y-prosemirror stores a mark under in a `Y.XmlText`'s formatting: its name, or for a mark
 * that may overlap its own kind (`excludes: ''`: a comment, and a deletion since ED-005 section
 * 12.26) `name--` and a hash of its JSON, so that two such marks are two attributes.
 * UPSTREAM-12: y-prosemirror 1.3.7 does not export its `hashOfJSON`, so it is repeated here (lib0's
 * SHA-256 over `encodeAny`, folded to six bytes, in base64); `store.test.mts` holds it to the
 * key y-prosemirror writes.
 */
export function yAttributeKey(mark: Mark): string {
  if (mark.type.excludes(mark.type)) return mark.type.name;
  const digest = sha256.digest(buf.encodeAny(mark.toJSON()));
  const N = 6;
  for (let i = N; i < digest.length; i++) digest[i % N] = digest[i % N]! ^ digest[i]!;
  return `${mark.type.name}--${buf.toBase64(digest.slice(0, N))}`;
}

/** What a resolution records, carried on its transaction. */
export interface ResolutionWrite {
  action: 'accept' | 'reject';
  ranges: ResolvedRange[];
}

interface ResolutionEntry extends ResolutionWrite {
  /** The state vector the resolver saw, encoded. */
  sv: Uint8Array;
}

/** The meta key a resolution's record travels under. */
export const RESOLUTION_WRITE = 'docx4j-resolution-write';
/** The origin of a repair's writes: not undone, like a peer's edit. */
export const REPAIR_ORIGIN = 'repair';
/** The origin of an acknowledgement. */
export const ACK_ORIGIN = 'repair-ack';

export function resolutionWriteOf(tr: Transaction): ResolutionWrite | undefined {
  return tr.getMeta(RESOLUTION_WRITE) as ResolutionWrite | undefined;
}

export function addResolutionWrite(tr: Transaction, write: ResolutionWrite): Transaction {
  return tr.setMeta(RESOLUTION_WRITE, write);
}

type Mapping = Parameters<typeof absolutePositionToRelativePosition>[2];

/**
 * The record of a resolution about to be made over `state`: each text piece of the changes, its
 * range anchored so that text typed at its edges lies inside. Through the view's live map when
 * given, else one rebuilt from the store. Undefined when nothing it resolves is text.
 */
export function resolutionRecord(fragment: Y.XmlFragment, schema: Schema, changes: readonly TrackedChangeInfo[], action: 'accept' | 'reject', live?: Mapping, state?: EditorState): ResolutionWrite | undefined {
  void state;
  const doc = fragment.doc!;
  let mapping = live;
  const ranges: ResolvedRange[] = [];
  const at = (pos: number, assoc: number): unknown => {
    mapping ??= initProseMirrorDoc(fragment, schema).mapping;
    const abs = Y.createAbsolutePositionFromRelativePosition(absolutePositionToRelativePosition(pos, fragment, mapping), doc);
    if (!abs) return null;
    return Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(abs.type, abs.index, assoc));
  };
  for (const c of changes) {
    for (const p of c.pieces) {
      if ((p.what !== 'text' && p.what !== 'format') || !p.mark) continue;
      const from = at(p.from, -1);
      const to = at(p.to, 0);
      if (from === null || to === null) continue;
      ranges.push({ from, to, mark: p.mark.type.name as RevisionKey, key: yAttributeKey(p.mark), attrs: { ...p.mark.attrs } });
    }
  }
  return ranges.length > 0 ? { action, ranges } : undefined;
}

/** Write a resolution's record: call it first inside the resolution's Yjs transaction, so the state vector is what the resolver saw. */
export function writeResolution(store: EditStore, write: ResolutionWrite): void {
  const sv = Y.encodeStateVector(store.ydoc);
  const id = `${store.ydoc.clientID}:${Y.getState(store.ydoc.store, store.ydoc.clientID)}`;
  const entry: ResolutionEntry = { sv, action: write.action, ranges: write.ranges };
  store.resolutions.set(id, entry);
}

/** A formatting write the repair makes. */
interface FormatOp {
  text: Y.XmlText;
  index: number;
  length: number;
  key: string;
  value: Record<string, unknown> | null;
}

/** Whether a resolution's effect on text it saw is a mark taken off (the ones a restoring marker can undo). */
function takesMarkOff(key: RevisionKey, action: 'accept' | 'reject'): boolean {
  return (key === 'insertion' && action === 'accept') || (key === 'deletion' && action === 'reject') || (key === 'formatChange' && action === 'accept');
}

/** The Y.XmlText segments between two absolute positions, in document order, a few blocks at most. */
function segments(a: Y.AbsolutePosition, b: Y.AbsolutePosition): [Y.XmlText, number, number][] {
  if (a.type === b.type) return a.type instanceof Y.XmlText ? [[a.type, a.index, b.index]] : [];
  const out: [Y.XmlText, number, number][] = [];
  if (a.type instanceof Y.XmlText) out.push([a.type, a.index, a.type.length]);
  // Onwards through the texts that follow, in document order, until the end's.
  let node: Y.AbstractType<any> | null = a.type;
  for (let guard = 0; guard < 64 && node; guard++) {
    node = nextInOrder(node);
    if (!node) break;
    if (node === b.type) { if (node instanceof Y.XmlText) out.push([node, 0, b.index]); break; }
    if (node instanceof Y.XmlText) out.push([node, 0, node.length]);
  }
  return out;
}

/** The type after `type` in a depth-first walk of the fragment: its first child, else its next sibling, else an ancestor's. */
function nextInOrder(type: Y.AbstractType<any>): Y.AbstractType<any> | null {
  const firstChild = (t: Y.AbstractType<any>): Y.AbstractType<any> | null => {
    for (let i = t._start; i; i = i.right) if (!i.deleted && i.content instanceof Y.ContentType) return i.content.type;
    return null;
  };
  const child = type instanceof Y.XmlElement ? firstChild(type) : null;
  if (child) return child;
  let t: Y.AbstractType<any> | null = type;
  while (t && t._item) {
    for (let i = t._item.right; i; i = i.right) if (!i.deleted && i.content instanceof Y.ContentType) return i.content.type;
    t = t._item.parent as Y.AbstractType<any> | null;
    if (t instanceof Y.XmlFragment) return null;
  }
  return null;
}

/**
 * The formatting writes that repair one range of one entry: scanned from the text's start, since
 * the attributes at a character are the markers before it. With `extend` (the range's last text),
 * on past its end over seen text still carrying the resolved mark: a restoring marker of the
 * typist's can reach there when Yjs's cleanup at the resolver removed the marker that ended the
 * change, and text touching the change with its attributes when it was resolved would have been
 * part of it (touching pieces of one type and author are one change), so such text gained them since.
 */
function repairRange(text: Y.XmlText, start: number, end: number, range: ResolvedRange, action: 'accept' | 'reject', seen: (client: number, clock: number) => boolean, concurrent: (client: number, clock: number) => boolean, out: FormatOp[], extend = false): void {
  const attrs = new Map<string, unknown>();
  let formatsSince: Y.Item[] = [];
  /** The run of concurrent text by one client the scan is in, and whether it had a marker of its own for the key. */
  let run: { client: number; explicit: boolean } | null = null;
  let index = 0;
  let pending: FormatOp | null = null;
  const push = (at: number, value: Record<string, unknown> | null): void => {
    if (pending && pending.index + pending.length === at && plainEqual(pending.value, value)) { pending.length++; return; }
    if (pending) out.push(pending);
    pending = { text, index: at, length: 1, key: range.key, value };
  };
  let extending = extend && takesMarkOff(range.mark, action);
  for (let item = text._start; item && (index < end || extending); item = item.right) {
    if (item.deleted) continue;
    const content = item.content;
    if (content instanceof Y.ContentFormat) {
      if (content.value === null) attrs.delete(content.key); else attrs.set(content.key, content.value);
      formatsSince.push(item);
      if (run && item.id.client === run.client && content.key === range.key) run.explicit = true;
      continue;
    }
    if (!item.countable) continue;
    if (content instanceof Y.ContentString) {
      const client = item.id.client;
      const ownMarker = formatsSince.some((f) => f.id.client === client && (f.content as Y.ContentFormat).key === range.key && !seen(f.id.client, f.id.clock));
      const isConcurrent = concurrent(client, item.id.clock);
      if (!isConcurrent) run = null;
      else if (!run || run.client !== client) run = { client, explicit: ownMarker };
      else if (ownMarker) run.explicit = true;
      for (let k = 0; k < item.length; k++) {
        const at = index + k;
        if (at < start) continue;
        const current = (attrs.get(range.key) ?? null) as Record<string, unknown> | null;
        const clock = item.id.clock + k;
        if (at >= end) {
          // Past the range: only the resolved mark's leak, while it lasts.
          if (extending && seen(client, clock) && current !== null && plainEqual(current, range.attrs)) { push(at, null); continue; }
          extending = false;
          break;
        }
        if (seen(client, clock)) {
          if (current !== null && takesMarkOff(range.mark, action) && plainEqual(current, range.attrs)) push(at, null);
        } else if (concurrent(client, clock) && range.mark === 'insertion' && current === null && run && !run.explicit) {
          push(at, range.attrs);
        }
      }
    } else {
      run = null;
    }
    formatsSince = [];
    index += item.length;
  }
  if (pending) out.push(pending);
}

/** The repair's formatting writes over the store as it is now: the same on every peer for the same state. */
export function repairWrites(store: EditStore): FormatOp[] {
  const doc = store.ydoc;
  const acks = new Map<string, Map<number, number>>();
  for (const [key, clock] of store.resolutionAcks.entries()) {
    const bar = key.lastIndexOf('|');
    const id = key.slice(0, bar);
    (acks.get(id) ?? acks.set(id, new Map()).get(id)!).set(Number(key.slice(bar + 1)), clock as number);
  }
  const out: FormatOp[] = [];
  for (const [id, value] of store.resolutions.entries()) {
    const entry = value as ResolutionEntry;
    const sv = Y.decodeStateVector(entry.sv);
    const resolver = Number(id.slice(0, id.indexOf(':')));
    const acked = acks.get(id) ?? new Map<number, number>();
    const seen = (client: number, clock: number): boolean => clock < (sv.get(client) ?? 0);
    const concurrent = (client: number, clock: number): boolean => client !== resolver && !seen(client, clock) && (!acked.has(client) || clock < acked.get(client)!);
    for (const range of entry.ranges) {
      const a = Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(range.from), doc);
      const b = Y.createAbsolutePositionFromRelativePosition(Y.createRelativePositionFromJSON(range.to), doc);
      if (!a || !b) continue;
      const parts = segments(a, b);
      parts.forEach(([text, start, end], i) => repairRange(text, start, end, range, entry.action, seen, concurrent, out, i === parts.length - 1));
    }
  }
  return out;
}

/**
 * The repair on one store (ED-005 section 12.15): after each remote update, the resolutions that
 * reached this peer for the first time acknowledged with its clock at arrival, then the repair's
 * writes, if any, made under `REPAIR_ORIGIN`. Outside a room there are no remote updates.
 */
export class RoomRepair {
  /** Entries that arrived and are not yet acknowledged, with this peer's clock when they did. */
  private readonly arrived = new Map<string, number>();
  private scheduled = false;
  private destroyed = false;
  /** How many repairs wrote something (tests). */
  repairs = 0;

  private readonly onAfter = (txn: Y.Transaction): void => {
    if (txn.local || this.destroyed) return;
    const me = this.store.ydoc.clientID;
    // This peer's clock now, before any later edit of its own: what it had typed when the resolution reached it.
    const clock = Y.getState(this.store.ydoc.store, me);
    for (const id of this.store.resolutions.keys()) {
      if (id.startsWith(`${me}:`) || this.arrived.has(id) || this.store.resolutionAcks.has(`${id}|${me}`)) continue;
      this.arrived.set(id, clock);
    }
    if (this.store.resolutions.size > 0) this.schedule();
  };

  constructor(private readonly store: EditStore) {
    store.ydoc.on('afterTransaction', this.onAfter);
  }

  private schedule(): void {
    if (this.scheduled) return;
    this.scheduled = true;
    queueMicrotask(() => { this.scheduled = false; if (!this.destroyed) this.run(); });
  }

  /** Acknowledge what arrived, then repair. Public for tests. */
  run(): void {
    const me = this.store.ydoc.clientID;
    if (this.arrived.size > 0) {
      const acks = [...this.arrived];
      this.arrived.clear();
      this.store.ydoc.transact(() => { for (const [id, clock] of acks) this.store.resolutionAcks.set(`${id}|${me}`, clock); }, ACK_ORIGIN);
    }
    const writes = repairWrites(this.store);
    if (writes.length === 0) return;
    this.repairs++;
    this.store.ydoc.transact(() => {
      for (const w of writes) w.text.format(w.index, w.length, { [w.key]: w.value });
    }, REPAIR_ORIGIN);
  }

  destroy(): void {
    this.destroyed = true;
    this.store.ydoc.off('afterTransaction', this.onAfter);
  }
}

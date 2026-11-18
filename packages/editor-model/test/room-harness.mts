// SPDX-License-Identifier: GPL-3.0-or-later
//
// Two stores linked in memory (ED-004 section 3.10): each store's updates applied to the other,
// which is all a provider does. Held back, the updates queue on both sides; released, they are
// delivered in an order the test names, so convergence is shown not to rest on one replay order
// (section 2.1's first review point).
import * as Y from 'yjs';
import { EditorDocument, newEditStore, type EditStore } from '../src/index.mjs';

export type ReleaseOrder = 'aFirst' | 'bFirst' | 'interleaved' | 'reversed';

export const RELEASE_ORDERS: readonly ReleaseOrder[] = ['aFirst', 'bFirst', 'interleaved', 'reversed'];

export class Link {
  private readonly toB: Uint8Array[] = [];
  private readonly toA: Uint8Array[] = [];
  private held = false;

  constructor(readonly a: Y.Doc, readonly b: Y.Doc) {
    a.on('update', (update: Uint8Array, origin: unknown) => { if (origin !== this) this.send(this.toB, this.b, update); });
    b.on('update', (update: Uint8Array, origin: unknown) => { if (origin !== this) this.send(this.toA, this.a, update); });
  }

  private send(queue: Uint8Array[], target: Y.Doc, update: Uint8Array): void {
    if (this.held) queue.push(update); else Y.applyUpdate(target, update, this);
  }

  /** From now on updates queue on both sides until `release`. */
  hold(): void {
    this.held = true;
  }

  /** Deliver the queued updates in the named order, then go on delivering at once. */
  release(order: ReleaseOrder): void {
    this.held = false;
    const toB = this.toB.splice(0);
    const toA = this.toA.splice(0);
    const apply = (target: Y.Doc, u: Uint8Array): void => Y.applyUpdate(target, u, this);
    switch (order) {
      case 'aFirst': toB.forEach((u) => apply(this.b, u)); toA.forEach((u) => apply(this.a, u)); break;
      case 'bFirst': toA.forEach((u) => apply(this.a, u)); toB.forEach((u) => apply(this.b, u)); break;
      case 'interleaved':
        for (let i = 0; i < Math.max(toA.length, toB.length); i++) { if (toB[i]) apply(this.b, toB[i]!); if (toA[i]) apply(this.a, toA[i]!); }
        break;
      case 'reversed': [...toB].reverse().forEach((u) => apply(this.b, u)); [...toA].reverse().forEach((u) => apply(this.a, u)); break;
    }
  }
}

export interface Room {
  creator: EditorDocument;
  joiner: EditorDocument;
  link: Link;
  base: Uint8Array;
}

/** A room over a fixture: the creator opens and shares, the joiner syncs and joins over the same base. */
export async function room(base: Uint8Array, name = 'fixture.docx', before?: (creator: EditorDocument) => Promise<void>): Promise<Room> {
  const creator = await EditorDocument.open(base);
  if (before) await before(creator);
  await creator.share(base, name);
  const store: EditStore = newEditStore();
  Y.applyUpdate(store.ydoc, Y.encodeStateAsUpdate(creator.store.ydoc));
  const link = new Link(creator.store.ydoc, store.ydoc);
  const joiner = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
  return { creator, joiner, link, base };
}

/** The ids of every node of a document, in order. */
export function nodeIds(doc: { descendants(f: (n: { attrs: Record<string, unknown> }) => boolean | void): void }): string[] {
  const out: string[] = [];
  doc.descendants((n) => { const id = n.attrs['sourceId']; if (typeof id === 'string') out.push(id); });
  return out;
}

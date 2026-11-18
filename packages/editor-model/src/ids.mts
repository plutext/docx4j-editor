// SPDX-License-Identifier: GPL-3.0-or-later
//
// Element identity for the projection (ED-002 section 3.4): every element that becomes a node
// or a `runSource` mark gets a session-local id, held in a WeakMap from the element and a Map
// back to it. Ids are never written to the file and never into the tree; they travel only in
// node attributes, which Yjs carries, so they survive a remote rebuild of a node.
//
// ED-004 sections 3.6 and 3.7: one registry, several allocators. The body's projection of the
// base file gives `e1`, `e2`, ... ; each story (a note, a header, a comment) allocates in a scope
// of its own (`footnote:3/e1`), so its ids depend on the base file and the story alone, never on
// the order a peer opened stories in. Every id minted after an allocator's base projection (a
// tree run's new elements) carries the peer's Yjs `clientID` (`c1234:e57`, `footnote:3/c1234:e2`),
// so no two peers, and no recovered session, ever mint the same id; such an element is always in
// the store's `sources` map (sources.mts). Lookups go through the one registry whatever the scope.
import type { Element } from '@docx4j/core-ts';

/** The ids and elements every allocator of a session shares. */
class Registry {
  readonly byElement = new WeakMap<object, string>();
  readonly byId = new Map<string, Element>();
  /** The minting token, `c<clientID>:`, once the session knows its Yjs document. */
  client: string | undefined;
  readonly scopes = new Map<string, SourceIds>();
  /** Ids handed out by `idOf` (not by `assign`), for the check that a replayed entry minted nothing. */
  minted = 0;
}

export class SourceIds {
  private next = 1;
  /** Set once the allocator's base projection is over: from then on its ids carry the client token. */
  private sealed = false;

  constructor(private readonly registry: Registry = new Registry(), private readonly prefix = '') {}

  /** The id of an element, assigned on first use. */
  idOf(element: Element): string {
    let id = this.registry.byElement.get(element);
    if (id === undefined) {
      id = `${this.prefix}${this.sealed && this.registry.client ? this.registry.client : ''}e${this.next++}`;
      this.registry.byElement.set(element, id);
      this.registry.byId.set(id, element);
      this.registry.minted++;
    }
    return id;
  }

  /** The id an element already has, or undefined. */
  existingIdOf(element: Element): string | undefined {
    return this.registry.byElement.get(element);
  }

  get(id: string): Element | undefined {
    return this.registry.byId.get(id);
  }

  get size(): number {
    return this.registry.byId.size;
  }

  /** How many ids `idOf` has minted in this registry so far. */
  get mintedCount(): number {
    return this.registry.minted;
  }

  /**
   * The allocator of a story (ED-004 section 3.6): the same registry, a counter and a prefix of
   * its own (`footnote:3/`), kept per prefix so that projecting the story again continues its
   * counter rather than restarting it.
   */
  scope(prefix: string): SourceIds {
    let scoped = this.registry.scopes.get(prefix);
    if (!scoped) {
      scoped = new SourceIds(this.registry, prefix);
      this.registry.scopes.set(prefix, scoped);
    }
    return scoped;
  }

  /**
   * The base projection is over: every id minted from now on carries the client token. The
   * token is set once per session (`c<clientID>:`); a story scope sealed later takes it too.
   */
  seal(client?: string): void {
    if (client !== undefined) this.registry.client = client;
    this.sealed = true;
  }

  /** Whether this allocator's base projection is over. */
  get isSealed(): boolean {
    return this.sealed;
  }

  /**
   * Give an element a known id (an entry of the `sources` map replayed, ED-004 section 3.7): the
   * element's previous id, if any, is dropped, and the id's previous element, if any, loses it.
   */
  assign(element: Element, id: string): void {
    const old = this.registry.byElement.get(element);
    if (old !== undefined && old !== id) this.registry.byId.delete(old);
    const previous = this.registry.byId.get(id);
    if (previous && previous !== element) this.registry.byElement.delete(previous);
    this.registry.byElement.set(element, id);
    this.registry.byId.set(id, element);
  }

  /** Drop an element's id, so the next projection mints a new one for it (a block changed in place by a tree run). */
  forget(element: Element): void {
    const id = this.registry.byElement.get(element);
    if (id === undefined) return;
    this.registry.byElement.delete(element);
    if (this.registry.byId.get(id) === element) this.registry.byId.delete(id);
  }
}

/**
 * A random integer in [min, max] not among `used` (ED-004 section 3.4: identifiers in a room are
 * random, as Word's own `w14:paraId` and control `w:id` are, so two peers never pick the same).
 */
export function randomId(used: ReadonlySet<number>, min = 2 ** 20, max = 2 ** 31 - 1): number {
  for (;;) {
    const n = min + Math.floor(Math.random() * (max - min + 1));
    if (!used.has(n)) return n;
  }
}

/** A relationship id for a room: `rId` and eight random base-36 characters. */
export function randomRelationshipId(): string {
  let s = '';
  while (s.length < 8) s += Math.random().toString(36).slice(2);
  return `rId${s.slice(0, 8)}`;
}

/** Whether an id is a base file's (the body's or a story's projection at open), not one minted since (`c<clientID>:`). */
export function isBaseId(id: string): boolean {
  return !/(^|\/)c\d+:/.test(id);
}

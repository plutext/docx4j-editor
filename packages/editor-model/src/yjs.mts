// SPDX-License-Identifier: GPL-3.0-or-later
//
// The edit layer's store (ED-002 section 3.5; ED-001 sections 2 and 6, Appendix B): one `Y.Doc`
// holding the body as a `Y.XmlFragment` and, from the first commit, the side maps E1.e and E2
// fill. The room schema (ED-004 section 3.2; ED-001 section 6): the maps below, a change to which
// is a schema version bump (`ROOM_SCHEMA`) that older peers refuse.
import * as Y from 'yjs';
import { Node as PMNode, type Mark, type Schema } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode, updateYFragment } from 'y-prosemirror';
import { plainEqual } from './json.mjs';
import { writeResolution, yAttributeKey, type ResolutionWrite } from './resolutions.mjs';

export const BODY_FRAGMENT = 'body';
export const FOOTNOTES_MAP = 'footnotes';
export const ENDNOTES_MAP = 'endnotes';
export const COMMENTS_MAP = 'comments';
export const HEADERS_MAP = 'headers';
export const CONTROLS_MAP = 'controls';
export const PACKAGE_MAP = 'package';
/** Every element that is not in the base file, as its XML and its ids by path (ED-004 section 3.7). */
export const SOURCES_MAP = 'sources';
/** The styles part's elements by id, the numbering part's by `num:` and `abstract:` key, other parts whole, relationships by source and id (ED-004 section 3.3). */
export const STYLES_MAP = 'styles';
export const NUMBERING_MAP = 'numbering';
export const PARTS_MAP = 'parts';
export const RELATIONSHIPS_MAP = 'relationships';
/** What a room is: its schema version, the base file, its hash and name (ED-004 sections 3.1, 4.4). */
export const ROOM_MAP = 'room';
/** A resolution's record in a room, by `<client>:<clock>` (ED-005 section 12.15, the room repair). */
export const RESOLUTIONS_MAP = 'resolutions';
/** When a resolution reached each peer, by `<entry>|<client>`: that client's clock then. */
export const RESOLUTION_ACKS_MAP = 'resolutionAcks';
/** A comment's thread data by its id: a `Y.Map` per comment, one field each (ED-005 section 7.2). */
export const COMMENT_THREADS_MAP = 'commentThreads';
/** An agent's Stop, by its name (ED-005 section 12.28; `agents.mts`). */
export const AGENTS_MAP = 'agents';
/** What the agents in the room did, a call per entry (`agents.mts`). */
export const AGENT_LOG_ARRAY = 'agentLog';
/**
 * The room schema's version: a peer refuses a room of any other version (ED-005 section 5.9). 2
 * since E4.c, whose fragments hold tracked changes as content where a room of version 1 has a
 * deletion as an atom; 3 since the room repair (section 12.15), whose records and
 * acknowledgements are two maps a room of version 2 has not; 4 since E4.e (section 7.2), whose
 * comments are edited through their fragments and a `commentThreads` map; 5 since core-ts 0.3.0
 * (section 12.26), whose fragments key a deletion by a hash of its attributes (two may overlap)
 * and hold dates read as Word means them; 6 since ED-009, whose fragments carry `alt` on the
 * `runSource` mark and the atoms (a peer without it would write a branch's text as the run's and
 * drop the element); 7 since core-ts 0.7.0 (ED-008 section 8.8 install item 18), whose reads of
 * some files project otherwise (`PROJECTION` 4); rooms are not migrated.
 */
export const ROOM_SCHEMA = 7;
/**
 * The projection a store's fragments were made by, in the package map under `projection`: 2 since
 * E4.c (ED-005 section 12.3); 3 since core-ts 0.3.0 (section 12.26), which reads a revision's
 * `w:date` as local time where there is no `w16du:dateUtc`, and whose deletions may overlap; 4 since
 * core-ts 0.7.0 with objects 0.8.0 (ED-008 section 8.8 install item 18), which read some files
 * otherwise (`rereadsOtherwise` in the session). A store without it, or with an older one, was made
 * before, and may hold the document in a shape this import no longer gives: it is refused for such a
 * document (`ProjectionVersionError`), and stamped for any other, whose projection is the same.
 */
export const PROJECTION = 4;
export const PROJECTION_KEY = 'projection';
/** The origin of the room map's writes: not undone, like a part edit. */
export const ROOM_ORIGIN = 'room';

/** A join before the room's state has arrived, or a story bound in a room before its fragment has (ED-004 sections 3.1, 3.9). */
export class RoomNotReadyError extends Error {
  override name = 'RoomNotReadyError';
}

/** A room made by another version of the editor, whose schema this one does not share (ED-004 section 3.1; ED-005 section 5.9). */
export class RoomVersionError extends Error {
  override name = 'RoomVersionError';
}

/** A persisted store made by an earlier projection, for a document it would misread (ED-005 section 12.3). */
export class ProjectionVersionError extends Error {
  override name = 'ProjectionVersionError';
}

/** The origin of the local user's transactions; `Y.UndoManager` (E1.b) tracks only these. */
export const LOCAL_ORIGIN = 'local';
/** The origin of the initial load and of re-imports after a tree mutation (decision 20). */
export const IMPORT_ORIGIN = 'import';

export interface EditStore {
  ydoc: Y.Doc;
  body: Y.XmlFragment;
  /** A fragment per footnote, keyed by the note's id (ED-003 section 6.2), from the first editable view over it. */
  footnotes: Y.Map<Y.XmlFragment>;
  /** The same for endnotes (section 9 item 22: a map of their own). */
  endnotes: Y.Map<Y.XmlFragment>;
  /** A fragment per comment, keyed by its id, from the first editable view over it or its making (ED-005 section 7.2). */
  comments: Y.Map<Y.XmlFragment>;
  /** A comment's thread data by its id (`ThreadFields`), for the comments made here and the file's that changed. */
  commentThreads: Y.Map<Y.Map<unknown>>;
  /** A fragment per header or footer part, keyed by part name (section 9 item 3). */
  headers: Y.Map<Y.XmlFragment>;
  /** Row- and cell-level control groups by `groupId` (ED-003 section 4.2, section 9 item 9): `form` and `sdtPr`. */
  controls: Y.Map<Y.Map<unknown>>;
  package: Y.Map<unknown>;
  /** `SourceEntry`s by the block's id (ED-004 section 3.7). */
  sources: Y.Map<unknown>;
  styles: Y.Map<unknown>;
  numbering: Y.Map<unknown>;
  parts: Y.Map<unknown>;
  relationships: Y.Map<unknown>;
  room: Y.Map<unknown>;
  /** Resolutions' records (the room repair, ED-005 section 12.15). */
  resolutions: Y.Map<unknown>;
  /** When each resolution reached each peer. */
  resolutionAcks: Y.Map<unknown>;
  /** An agent's state in the room by its name: whether a person stopped it (`agents.mts`). */
  agents: Y.Map<unknown>;
  /** The agents' calls on the room (`AgentLogEntry`s), oldest first. */
  agentLog: Y.Array<unknown>;
}

export function newEditStore(ydoc: Y.Doc = new Y.Doc()): EditStore {
  return {
    ydoc,
    body: ydoc.getXmlFragment(BODY_FRAGMENT),
    footnotes: ydoc.getMap(FOOTNOTES_MAP),
    endnotes: ydoc.getMap(ENDNOTES_MAP),
    comments: ydoc.getMap(COMMENTS_MAP),
    commentThreads: ydoc.getMap(COMMENT_THREADS_MAP),
    headers: ydoc.getMap(HEADERS_MAP),
    controls: ydoc.getMap(CONTROLS_MAP),
    package: ydoc.getMap(PACKAGE_MAP),
    sources: ydoc.getMap(SOURCES_MAP),
    styles: ydoc.getMap(STYLES_MAP),
    numbering: ydoc.getMap(NUMBERING_MAP),
    parts: ydoc.getMap(PARTS_MAP),
    relationships: ydoc.getMap(RELATIONSHIPS_MAP),
    room: ydoc.getMap(ROOM_MAP),
    resolutions: ydoc.getMap(RESOLUTIONS_MAP),
    resolutionAcks: ydoc.getMap(RESOLUTION_ACKS_MAP),
    agents: ydoc.getMap(AGENTS_MAP),
    agentLog: ydoc.getArray(AGENT_LOG_ARRAY),
  };
}

// ---------------------------------------------------------------- comment threads (E4.e)

/**
 * A comment's thread data as the `commentThreads` map holds it (ED-005 section 7.2): every field
 * for a comment made in the editor, and only what changed for one from the file, whose other
 * fields are the file's. Each field is its own entry, so two peers resolving and deleting at once
 * both hold.
 */
export interface ThreadFields {
  author?: string;
  initials?: string;
  /** ISO 8601, as made. */
  date?: string;
  /** The id of the comment this one replies to; null for a thread's first. */
  parent?: string | null;
  done?: boolean;
  deleted?: boolean;
}

const THREAD_FIELDS = ['author', 'initials', 'date', 'parent', 'done', 'deleted'] as const;

/** A comment's entry, or undefined when the map has none. */
export function readThreadEntry(store: EditStore, id: string): ThreadFields | undefined {
  const map = store.commentThreads.get(id);
  if (!map) return undefined;
  const out: ThreadFields = {};
  for (const k of THREAD_FIELDS) if (map.has(k)) (out as Record<string, unknown>)[k] = map.get(k);
  return out;
}

/** Set fields of a comment's entry, the entry made when it has none; call inside the transaction that should carry it. */
export function writeThreadFields(store: EditStore, id: string, fields: ThreadFields): void {
  let map = store.commentThreads.get(id);
  if (!map) { map = new Y.Map<unknown>(); store.commentThreads.set(id, map); }
  for (const k of THREAD_FIELDS) if (k in fields) map.set(k, (fields as Record<string, unknown>)[k]);
}

/** A write a comment command asks for, carried on its transaction's meta as control writes are (ED-005 section 7.3). */
export interface CommentWrite {
  id: string;
  fields: ThreadFields;
}

export const COMMENT_WRITES = 'docx4j-comment-writes';

export function commentWritesOf(tr: Transaction): CommentWrite[] {
  return (tr.getMeta(COMMENT_WRITES) as CommentWrite[] | undefined) ?? [];
}

export function addCommentWrite(tr: Transaction, write: CommentWrite): Transaction {
  return tr.setMeta(COMMENT_WRITES, [...commentWritesOf(tr), write]);
}

// ---------------------------------------------------------------- story fragments (E2.d)

/**
 * A story's fragment in a side map: the one the map holds, else a new one loaded with `doc`
 * under the import origin (ED-003 section 6.2: not an undo step, not a dirty mark). The fragment
 * is put in the map before it is filled, since y-prosemirror fills only an integrated fragment.
 */
export function fragmentIn(store: EditStore, map: Y.Map<Y.XmlFragment>, key: string, doc: () => PMNode, create = true): Y.XmlFragment {
  const existing = map.get(key);
  if (existing) return existing;
  // In a room only the creator (at Share) and the command that makes a story create a fragment;
  // a peer binding a story whose fragment has not arrived would make an orphan (ED-004 section 3.9).
  if (!create) throw new RoomNotReadyError(`the story ${key} has no fragment in the room yet`);
  const fragment = new Y.XmlFragment();
  store.ydoc.transact(() => {
    map.set(key, fragment);
    prosemirrorToYXmlFragment(doc(), fragment);
  }, IMPORT_ORIGIN);
  return fragment;
}

/** A story fragment as a document of the schema (the root carries no attributes). */
export function readFragment(fragment: Y.XmlFragment, schema: Schema): PMNode {
  return yXmlFragmentToProseMirrorRootNode(fragment, schema);
}

/** Replace a story's content as one transaction, diffing as `replaceDocument` does. */
export function replaceFragment(store: EditStore, fragment: Y.XmlFragment, doc: PMNode, origin: unknown = LOCAL_ORIGIN): void {
  store.ydoc.transact(() => {
    updateYFragment(store.ydoc, fragment, doc, { mapping: new Map(), isOMark: new Map() });
    // UPSTREAM-15: the formatting its diff misjudges.
    syncTextFormatting(fragment, doc);
  }, origin);
}

/**
 * The body's final `w:sectPr` (ED-001 section 3.1: `doc.attrs.finalSectPr`) lives in the
 * `package` map under this key: a `Y.XmlFragment` holds the root node's children only, never
 * its attributes, so the attribute is a projection of the map entry, written on load and
 * replace and re-applied on read. Section commands (E1.b) write the map.
 */
export const FINAL_SECT_PR = 'finalSectPr';

/** Load a ProseMirror document into the body fragment (first load, or a re-import). */
export function loadDocument(store: EditStore, doc: PMNode, origin: unknown = IMPORT_ORIGIN): void {
  store.ydoc.transact(() => {
    if (store.body.length > 0) store.body.delete(0, store.body.length);
    prosemirrorToYXmlFragment(doc, store.body);
    setFinalSectPr(store, doc.attrs['finalSectPr']);
    if (store.package.get(PROJECTION_KEY) !== PROJECTION) store.package.set(PROJECTION_KEY, PROJECTION);
  }, origin);
}

function setFinalSectPr(store: EditStore, value: unknown): void {
  // Only on a change: a Yjs `set` is always a new item, and it would make every replace an undo
  // step that touches the map.
  const current = store.package.get(FINAL_SECT_PR) ?? null;
  if (plainEqual(current, value ?? null)) return;
  if (value === null || value === undefined) store.package.delete(FINAL_SECT_PR);
  else store.package.set(FINAL_SECT_PR, value);
}

/** The body fragment as a ProseMirror document (headless use; the view binds with `ySyncPlugin`). */
export function readDocument(store: EditStore, schema: Schema): PMNode {
  const fromFragment = yXmlFragmentToProseMirrorRootNode(store.body, schema);
  return schema.topNodeType.create({ finalSectPr: store.package.get(FINAL_SECT_PR) ?? null }, fromFragment.content);
}

/**
 * Replace the body with a new ProseMirror document as one transaction, diffing against what the
 * fragment holds (y-prosemirror's `updateYFragment`), so that unchanged nodes keep their Yjs
 * identity and a `Y.UndoManager` sees one step. Control-map writes a command asked for travel
 * in the same transaction, so a wrapped row group and its properties are one undo step.
 */
export function replaceDocument(store: EditStore, doc: PMNode, origin: unknown = LOCAL_ORIGIN, writes: readonly StoreWrite[] = [], resolution?: ResolutionWrite, before?: PMNode): void {
  store.ydoc.transact(() => {
    // First, so that its state vector is what the resolver saw (ED-005 section 12.15).
    if (resolution) writeResolution(store, resolution);
    // UPSTREAM-16: the store's children made to match the new document's in kind first, so that the
    // diff never replaces a child with a copy of its neighbour.
    if (before && store.body.length === before.childCount) {
      const stretch = changedStretch(before, doc);
      if (stretch) alignBlocks(store.body, doc, stretch.from, stretch.toDoc, stretch.toBefore, (i) => before.child(i));
    }
    updateYFragment(store.ydoc, store.body, doc, { mapping: new Map(), isOMark: new Map() });
    // UPSTREAM-15: the formatting its diff misjudges (bold taken off as a formatChange comes).
    syncTextFormatting(store.body, doc, before);
    setFinalSectPr(store, doc.attrs['finalSectPr']);
    for (const w of writes) writeStoreWrite(store, w);
  }, origin);
}

/** The top-level blocks that differ between two documents, by identity from each end: `[from, toDoc)` in `doc`, `[from, toBefore)` in `before`; none when they are the same. */
function changedStretch(before: PMNode, doc: PMNode): { from: number; toDoc: number; toBefore: number } | undefined {
  const n = before.childCount;
  const m = doc.childCount;
  let from = 0;
  while (from < n && from < m && before.child(from) === doc.child(from)) from++;
  let toBefore = n;
  let toDoc = m;
  while (toBefore > from && toDoc > from && before.child(toBefore - 1) === doc.child(toDoc - 1)) { toBefore--; toDoc--; }
  return from === toBefore && from === toDoc ? undefined : { from, toDoc, toBefore };
}

/** A child's kind as y-prosemirror pairs children: an element's node name, or text. */
const TEXT_KIND = '#text';

/** Children as y-prosemirror sees them (adjacent text nodes one `Y.XmlText`): each one's kind, its kind and text, and its node (none for text). */
interface Items { kinds: string[]; keys: string[]; nodes: (PMNode | undefined)[] }

function itemsOf(children: readonly (PMNode | undefined)[]): Items {
  const out: Items = { kinds: [], keys: [], nodes: [] };
  children.forEach((c, i) => {
    if (c?.isText) {
      const last = out.kinds.length - 1;
      if (out.kinds[last] === TEXT_KIND) { out.keys[last] += c.text!; return; }
      out.kinds.push(TEXT_KIND); out.keys.push(`${TEXT_KIND}\u0000${c.text!}`); out.nodes.push(undefined);
    } else if (c) {
      out.kinds.push(c.type.name); out.keys.push(`${c.type.name}\u0000${c.textContent}`); out.nodes.push(c);
    } else {
      // A child whose node is not known: it matches by kind alone.
      out.kinds.push(''); out.keys.push(`\u0001${i}`); out.nodes.push(undefined);
    }
  });
  return out;
}

function childrenOf(node: PMNode, from = 0, to = node.childCount): PMNode[] {
  const out: PMNode[] = [];
  for (let i = from; i < to; i++) out.push(node.child(i));
  return out;
}

function yKind(y: unknown): string {
  return y instanceof Y.XmlText ? TEXT_KIND : y instanceof Y.XmlElement ? y.nodeName : '';
}

const KEEP = 0;
const DELETE = 1;
const INSERT = 2;

/**
 * Myers' shortest edit script from `a` to `b`: KEEP, DELETE (an element of `a`), INSERT (one of `b`),
 * in order; undefined when they are more than `maxD` edits apart.
 */
export function editScript(a: readonly string[], b: readonly string[], maxD = 1000): number[] | undefined {
  const n = a.length;
  const m = b.length;
  const max = n + m;
  const off = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let found = -1;
  for (let d = 0; d <= Math.min(max, maxD) && found < 0; d++) {
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[off + k - 1]! < v[off + k + 1]!) ? v[off + k + 1]! : v[off + k - 1]! + 1;
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) { x++; y++; }
      v[off + k] = x;
      if (x >= n && y >= m) { found = d; break; }
    }
    trace.push(v.slice(off - d, off + d + 1));
  }
  if (found < 0) return undefined;
  const ops: number[] = [];
  let x = n;
  let y = m;
  for (let d = found; d > 0; d--) {
    const prev = trace[d - 1]!;
    const at = (k: number): number => prev[k + d - 1]!;
    const k = x - y;
    const prevK = k === -d || (k !== d && at(k - 1) < at(k + 1)) ? k + 1 : k - 1;
    const prevX = at(prevK);
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) { ops.push(KEEP); x--; y--; }
    ops.push(x === prevX ? INSERT : DELETE);
    x = prevX;
    y = prevY;
  }
  while (x > 0 && y > 0) { ops.push(KEEP); x--; y--; }
  return ops.reverse();
}

/**
 * How the children `had` become `want`: first the children of the same kind and text, then, between
 * them, the children of the same kind. Undefined when too far apart to say.
 */
function pairing(had: Items, want: Items): number[] | undefined {
  const same = had.kinds.length === want.kinds.length && had.kinds.every((k, i) => k === want.kinds[i]);
  if (same) return had.kinds.map(() => KEEP);
  const byKind = (i0: number, i1: number, j0: number, j1: number): number[] =>
    editScript(had.kinds.slice(i0, i1), want.kinds.slice(j0, j1)) ?? [...Array<number>(i1 - i0).fill(DELETE), ...Array<number>(j1 - j0).fill(INSERT)];
  const coarse = editScript(had.keys, want.keys);
  if (!coarse) return editScript(had.kinds, want.kinds);
  const ops: number[] = [];
  let i = 0;
  let j = 0;
  let k = 0;
  while (k < coarse.length) {
    if (coarse[k] === KEEP) { ops.push(KEEP); i++; j++; k++; continue; }
    const [i0, j0] = [i, j];
    while (k < coarse.length && coarse[k] !== KEEP) { if (coarse[k] === DELETE) i++; else j++; k++; }
    ops.push(...byKind(i0, i, j0, j));
  }
  return ops;
}

/**
 * UPSTREAM-16: y-prosemirror's `updateYFragment` pairs a parent's children from each end, and where
 * the child it must remove is not of the kind of the new child in its place (a move's range marker
 * beside a kept paragraph; an inline atom before a paragraph's text), it deletes the child and inserts
 * a fresh copy of the new one, whose own old element or text is deleted after; where it pairs two
 * paragraphs of different text, it rewrites the one's text into the other's. One writer sees nothing
 * of it. Two peers making the same change at once (two people accepting the same move, both pressing
 * Accept all) each insert a copy, and both copies survive the merge: a paragraph twice, a text twice
 * over in its paragraph, or its words twice between a common start and end. So, first, each changed
 * parent's children are matched to the new node's: by kind and text, then by kind between those, by
 * the shortest edit script; the children the new node has no place for are deleted, an empty element
 * or text is put where a new child goes, and the children kept are matched the same way, down the
 * tree. The diff then meets only children of the kind it is updating, each against its own match,
 * which it updates in place. `had` is what the old document held for the children `[yFrom, yTo)`,
 * when it is known and agrees with them in kind.
 */
function alignChildren(parent: Y.XmlFragment | Y.XmlElement, yFrom: number, yTo: number, had: Items | undefined, want: Items, mapping?: Map<unknown, unknown>): void {
  const ys = parent.toArray().slice(yFrom, yTo);
  const kinds = ys.map(yKind);
  const old = had && had.kinds.length === kinds.length && had.kinds.every((k, i) => k === '' || k === kinds[i])
    ? { ...had, kinds }
    : { kinds, keys: kinds.map((_, i) => `\u0001${i}`), nodes: kinds.map(() => undefined) };
  const ops = pairing(old, want);
  if (!ops) return;
  const kept: [Y.XmlElement, PMNode | undefined, PMNode][] = [];
  let at = yFrom;
  let i = 0;
  let j = 0;
  for (const op of ops) {
    if (op === KEEP) {
      const y = ys[i];
      const p = want.nodes[j];
      if (y instanceof Y.XmlElement && p) kept.push([y, old.nodes[i], p]);
      at++; i++; j++;
    } else if (op === DELETE) {
      mapping?.delete(ys[i]);
      parent.delete(at, 1);
      i++;
    } else {
      parent.insert(at, [want.kinds[j] === TEXT_KIND ? new Y.XmlText() : new Y.XmlElement(want.kinds[j]!)]);
      at++; j++;
    }
  }
  for (const [y, before, p] of kept) {
    if (before === p) continue;
    alignChildren(y, 0, y.length, before ? itemsOf(childrenOf(before)) : undefined, itemsOf(childrenOf(p)), mapping);
  }
}

/**
 * UPSTREAM-16 at the top level: the fragment's blocks `[from, toY)` against the document's
 * `[from, toDoc)`, as `alignChildren` matches a parent's children, except that a block whose node is
 * the new document's own is kept as it is and never walked, so that only what changed is. `known(i)`
 * is the node the fragment's child `i` held before the write: the old document's block, or the one
 * the binding's mapping holds. Call it inside the write's transaction, before `updateYFragment`.
 */
export function alignBlocks(fragment: Y.XmlFragment, doc: PMNode, from: number, toDoc: number, toY: number, known: (i: number) => unknown, mapping?: Map<unknown, unknown>): void {
  if (toY > fragment.length) return;
  const knownAt = (i: number): PMNode | undefined => { const n = known(i); return n instanceof Object && 'type' in n ? n as PMNode : undefined; };
  // The blocks both sides share, in order: the anchors between which the rest is matched.
  const index = new Map<PMNode, number>();
  for (let j = from; j < toDoc; j++) index.set(doc.child(j), j);
  const anchors: [number, number][] = [];
  let last = from - 1;
  for (let i = from; i < toY; i++) {
    const node = knownAt(i);
    const j = node ? index.get(node) : undefined;
    if (j !== undefined && j > last) { anchors.push([i, j]); last = j; }
  }
  // From the last gap to the first, so that the earlier ones keep their places.
  const starts: [number, number][] = [[from, from], ...anchors.map(([i, j]): [number, number] => [i + 1, j + 1])];
  const ends: [number, number][] = [...anchors, [toY, toDoc]];
  for (let g = starts.length - 1; g >= 0; g--) {
    const [yStart, pStart] = starts[g]!;
    const [yEnd, pEnd] = ends[g]!;
    if (yStart === yEnd && pStart === pEnd) continue;
    const had: (PMNode | undefined)[] = [];
    for (let i = yStart; i < yEnd; i++) had.push(knownAt(i));
    alignChildren(fragment, yStart, yEnd, itemsOf(had), itemsOf(childrenOf(doc, pStart, pEnd)), mapping);
  }
}

/**
 * The text formatting y-prosemirror's diff misjudged, written (UPSTREAM-15). Its `equalYTextPText`
 * takes a text as unchanged when it has as many formatting attributes as the text has marks and
 * each of its attributes matches the mark of that name, a mark the text no longer has matching an
 * attribute with no values: so a text whose `bold`, `italic` or `rtl` (marks with no attributes)
 * went as another mark came keeps its old formatting in the store. With tracking on that is the
 * commonest formatting change there is: bold taken off text, which gains a `formatChange` for it.
 * This compares each text of the top-level blocks in scope with its marks, by the keys y-prosemirror
 * writes (`yAttributeKey`), and formats what differs. The scope is the blocks that differ from a
 * document the store held before, or a range of the document's top-level indexes; every block
 * without one. Call it after `updateYFragment`, inside the same transaction. Answers how many texts
 * it wrote.
 */
export function syncTextFormatting(fragment: Y.XmlFragment, doc: PMNode, scope?: PMNode | { from: number; to: number }): number {
  const children = fragment.toArray();
  let from = 0;
  let to = doc.childCount;
  if (scope instanceof PMNode) {
    const start = doc.content.findDiffStart(scope.content);
    if (start === null) return 0;
    const end = doc.content.findDiffEnd(scope.content)!;
    from = childIndexAt(doc, start);
    to = Math.max(from + 1, childIndexAt(doc, Math.max(start, end.a - 1)) + 1);
  } else if (scope) {
    ({ from, to } = scope);
  }
  let written = 0;
  for (let i = from; i < Math.min(to, children.length); i++) written += syncNode(children[i]!, doc.child(i));
  return written;
}

/** The index of the top-level child a position of the document's content is in (or before). */
function childIndexAt(doc: PMNode, pos: number): number {
  let at = 0;
  for (let i = 0; i < doc.childCount; i++) {
    at += doc.child(i).nodeSize;
    if (pos < at) return i;
  }
  return doc.childCount;
}

function syncNode(y: Y.XmlElement | Y.XmlText | Y.XmlHook, node: PMNode): number {
  if (!(y instanceof Y.XmlElement) || y.nodeName !== node.type.name || node.childCount === 0) return 0;
  // y-prosemirror's grouping: a run of text nodes is one `Y.XmlText`, any other node one element.
  const groups: (PMNode | PMNode[])[] = [];
  node.forEach((child) => {
    const last = groups[groups.length - 1];
    if (child.isText && Array.isArray(last)) last.push(child);
    else groups.push(child.isText ? [child] : child);
  });
  const ys = y.toArray();
  if (ys.length !== groups.length) return 0;
  let written = 0;
  groups.forEach((group, i) => {
    const yc = ys[i]!;
    if (Array.isArray(group)) { if (yc instanceof Y.XmlText) written += syncText(yc, group); }
    else written += syncNode(yc, group);
  });
  return written;
}

function syncText(ytext: Y.XmlText, texts: PMNode[]): number {
  const delta = ytext.toDelta() as { insert: unknown; attributes?: Record<string, unknown> }[];
  const wanted = texts.map((t) => ({ text: t.text!, attributes: attributesOf(t.marks) }));
  if (delta.map((d) => d.insert).join('') !== wanted.map((w) => w.text).join('')) return 0;
  if (sameFormatting(delta, wanted)) return 0;
  const cleared: Record<string, null> = {};
  for (const d of delta) for (const k of Object.keys(d.attributes ?? {})) cleared[k] = null;
  ytext.applyDelta(wanted.map((w) => ({ retain: w.text.length, attributes: { ...cleared, ...w.attributes } })));
  return 1;
}

function attributesOf(marks: readonly Mark[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const m of marks) if (m.type.name !== 'ychange') out[yAttributeKey(m)] = m.attrs;
  return out;
}

/** The formatting of each character alike, attribute values compared as y-prosemirror compares them (a null field is no field). */
function sameFormatting(delta: { insert: unknown; attributes?: Record<string, unknown> }[], wanted: { text: string; attributes: Record<string, unknown> }[]): boolean {
  let i = 0;
  let j = 0;
  let di = 0;
  let wj = 0;
  while (i < delta.length && j < wanted.length) {
    const d = delta[i]!;
    const w = wanted[j]!;
    if (!sameAttributes(d.attributes ?? {}, w.attributes)) return false;
    const dl = (d.insert as string).length - di;
    const wl = w.text.length - wj;
    const n = Math.min(dl, wl);
    di += n;
    wj += n;
    if (di === (d.insert as string).length) { i++; di = 0; }
    if (wj === w.text.length) { j++; wj = 0; }
  }
  return true;
}

function sameAttributes(a: Record<string, unknown>, b: Record<string, unknown>): boolean {
  const ka = Object.keys(a).filter((k) => a[k] !== null && a[k] !== undefined);
  const kb = Object.keys(b).filter((k) => b[k] !== null && b[k] !== undefined);
  return ka.length === kb.length && ka.every((k) => k in b && sameValue(a[k], b[k]));
}

function sameValue(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  return sameAttributes(a as Record<string, unknown>, b as Record<string, unknown>);
}

/** A side-map write a command's transaction carries: a control group's entry, or a comment's thread fields. */
export type StoreWrite = ControlWrite | CommentWrite;

/** Every side-map write a transaction carries, controls first. */
export function storeWritesOf(tr: Transaction): StoreWrite[] {
  return [...controlWritesOf(tr), ...commentWritesOf(tr)];
}

/** Apply one write; call inside the transaction that carries the document. */
export function writeStoreWrite(store: EditStore, w: StoreWrite): void {
  if ('groupId' in w) writeControl(store, w.groupId, w.entry);
  else writeThreadFields(store, w.id, w.fields);
}

// ---------------------------------------------------------------- the controls map (E2.b)

/** A row- or cell-level control group's entry: which form, and its `w:sdtPr` as plain JSON. */
export interface ControlEntry {
  form: 'row' | 'cell';
  sdtPr: unknown | null;
}

export function readControlEntry(store: EditStore, groupId: string): ControlEntry | undefined {
  const map = store.controls.get(groupId);
  if (!map) return undefined;
  const form = map.get('form');
  return { form: form === 'cell' ? 'cell' : 'row', sdtPr: map.get('sdtPr') ?? null };
}

/** Write an entry (null deletes it); only on a change, so an unchanged entry costs no update. Call inside a transaction. */
export function writeControl(store: EditStore, groupId: string, entry: ControlEntry | null): void {
  const current = readControlEntry(store, groupId);
  if (entry === null) { if (current) store.controls.delete(groupId); return; }
  if (current && current.form === entry.form && plainEqual(current.sdtPr, entry.sdtPr)) return;
  const map = new Y.Map<unknown>();
  map.set('form', entry.form);
  if (entry.sdtPr !== null) map.set('sdtPr', entry.sdtPr);
  store.controls.set(groupId, map);
}

/** Every entry, by group id. */
export function readControlEntries(store: EditStore): Map<string, ControlEntry> {
  const out = new Map<string, ControlEntry>();
  for (const id of store.controls.keys()) {
    const entry = readControlEntry(store, id);
    if (entry) out.set(id, entry);
  }
  return out;
}

/** A write a command asks for, carried on its transaction's meta. */
export interface ControlWrite {
  groupId: string;
  entry: ControlEntry | null;
}

/** The meta key the control writes travel under. */
export const CONTROL_WRITES = 'docx4j-control-writes';

export function controlWritesOf(tr: Transaction): ControlWrite[] {
  return (tr.getMeta(CONTROL_WRITES) as ControlWrite[] | undefined) ?? [];
}

export function addControlWrite(tr: Transaction, write: ControlWrite): Transaction {
  return tr.setMeta(CONTROL_WRITES, [...controlWritesOf(tr), write]);
}

/**
 * Apply a transaction's control writes to the store under an origin (the view's plugin does
 * this in the same Yjs transaction as the sync plugin's update; the headless runner passes them
 * to `replaceDocument`).
 */
export function applyControlWrites(store: EditStore, writes: readonly ControlWrite[], origin: unknown = LOCAL_ORIGIN): void {
  if (writes.length === 0) return;
  store.ydoc.transact(() => { for (const w of writes) writeControl(store, w.groupId, w.entry); }, origin);
}

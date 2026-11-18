// SPDX-License-Identifier: GPL-3.0-or-later
//
// The other stories of a document (ED-002 section 7.1; ED-001 sections 3.2, 3.3 and Appendix
// D.2; ED-003 section 6.2): footnotes and endnotes, comments with their threads, headers and
// footers with Word's rules for which one a page shows. Each part is read privately
// (`readContents()`) on first use and never at load (decision 5), and projected with the body's
// schema and walk into a document of its own. A note or a header enters Yjs when an editable
// view first binds to it (`fragmentOf`, under the import origin); a save writes a part only
// where a fragment's content differs from the story's baseline, removes a note whose reference
// the body lost, and appends a note the session made (section 9 item 21). Comments are edited
// the same way since E4.e (ED-005 section 7): a fragment per comment in the `comments` map, their
// thread data in `commentThreads`.
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import type * as Y from 'yjs';
import { Namespaces, typeNameOf, revisionDateOf, type MainDocumentPart, deepCopy, linkParents, runOf, p, t, marshalString, unmarshalString, markersOf, Body, Comment, CommentsPart, addCommentEntry, type CommentParts, type HeaderPart, type FooterPart, type FootnotesPart, type EndnotesPart, type Element } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as w15 from '@docx4j/generated-objects-ts/modules/org_docx4j_w15';
import { projectBlocks, newContext, type ProjectionContext } from './import.mjs';
import { exportBlocks, type RevisionIds } from './export.mjs';
import { fragmentIn, readFragment, replaceFragment, readThreadEntry, writeThreadFields, LOCAL_ORIGIN, IMPORT_ORIGIN } from './yjs.mjs';
import { commentElement, commentAnchors, initialsOf } from './comments.mjs';
import { compareProjections, type SelfCheckResult } from './selfcheck.mjs';
import { randomId } from './ids.mjs';
import { pathIds, assignPathIds, type SourceEntry } from './sources.mjs';

/** The `sources` key of a note made in a session (ED-004 section 3.7: an element not in the base file is always in `sources`). */
export const noteSourceKey = (kind: NoteKind, id: number): string => `note:${kind}:${id}`;

/** The `sources` key of a comment made in a session, as `noteSourceKey` (ED-005 section 7.2). */
export const commentSourceKey = (id: string): string => `comment:${id}`;

/** The ids a save writes for the comments made in a room: the comment's id in the room to its id in the file (ED-005 section 7.4). */
export type CommentIds = Map<number, number>;

/** The ids a save writes for room-made notes, by kind: the note's id in the room to its id in the file (ED-004 section 10.8 item 6). */
export type NoteIds = Map<NoteKind, Map<number, number>>;
import type { EditorDocument } from './session.mjs';

export interface Story {
  /** `footnote:3`, `endnote:2`, `comment:5`, `header:/word/header1.xml`: the story's name, and its fragment's key where it has one. */
  key: string;
  /** The projection as imported: the baseline a save compares against. The current document is `Stories.currentDoc`. */
  doc: PMNode;
  ctx: ProjectionContext;
}

export type NoteKind = 'footnote' | 'endnote';

/** The style ids a note needs (Word's built-ins, docx4j's defaults): the note's paragraph and its reference mark. */
export const NOTE_STYLES: Record<NoteKind, { text: string; reference: string }> = {
  footnote: { text: 'FootnoteText', reference: 'FootnoteReference' },
  endnote: { text: 'EndnoteText', reference: 'EndnoteReference' },
};

export interface Note extends Story {
  id: number;
  kind: NoteKind;
  /** No reference in the body pointed at it when the document was opened: it is never removed by a save. */
  orphan: boolean;
  /** The `w:footnote` or `w:endnote`; in the part, or, for a note this session made, waiting for a save that finds it referenced. */
  element: wml.CTFtnEdn;
  /** Made in this session and not yet in the part. */
  isNew: boolean;
}

/**
 * A comment as the session holds it (ED-005 section 7): its story, its element and its thread
 * data as the file gives it (or as it was made). Stable while the session lives; the listing
 * (`CommentThread`) is made from it afresh with the `commentThreads` entry laid over.
 */
export interface CommentRecord extends Story {
  id: string;
  /** The `w:comment`: the part's, or, for one made in this session, waiting for a save. */
  element: wml.Comments.Comment;
  /** Made in this session and not yet in the part. */
  isNew: boolean;
  /** No mark or reference in the body anchored it when the document opened: a save keeps it. */
  orphan: boolean;
  base: { author: string; initials: string; date: Date | undefined; parentId: string | null; done: boolean };
}

/** A comment as a panel lists it now: its record with the thread entry laid over, its replies under it. */
export interface CommentThread extends Story {
  id: string;
  author: string;
  initials: string;
  date: Date | undefined;
  parentId: string | null;
  done: boolean;
  /** No mark or reference in the body anchored it when the document opened. */
  orphan: boolean;
  isNew: boolean;
  replies: CommentThread[];
  record: CommentRecord;
}

/** What the comments part gave: the part and its tree when the document has one, and every comment by id. */
interface CommentsInfo {
  tree: wml.Comments | undefined;
  records: Map<string, CommentRecord>;
}

/** A comment's thread data for `Stories.newComment`. */
export interface NewComment {
  author: string;
  initials?: string;
  date?: Date;
  /** The comment it replies to. */
  parent?: string | null;
  text?: string;
}

export interface HeaderFooter extends Story {
  rId: string;
  kind: 'header' | 'footer';
  part: HeaderPart | FooterPart;
  /** The part's root (`w:hdr` or `w:ftr`), whose content a save rewrites. */
  element: wml.Hdr | wml.Ftr;
}

export type HeaderVariant = 'default' | 'first' | 'even';

export interface PageStories {
  header?: HeaderFooter;
  footer?: HeaderFooter;
  variant: HeaderVariant;
}

/** What a notes part gave: the part, its tree (private until a save writes it), and its notes by id. */
interface NotesInfo {
  part: FootnotesPart | EndnotesPart;
  tree: wml.CTFootnotes | wml.CTEndnotes;
  /** In the E1.e order: referenced ones by the baseline's reference order, then orphans; new ones appended. */
  notes: Note[];
  byId: Map<number, Note>;
}

/**
 * Project block content into a document of its own, in the story's own id scope (ED-004 section
 * 3.6): `footnote:3/e1`, ..., the same on every peer and in every session whatever order the
 * stories are opened in, in the session's one registry so every lookup still finds them. The
 * scope is sealed after its first projection, so what a later one mints carries the client.
 */
function projectStory(session: EditorDocument, content: Element[] | undefined, key: string): Story {
  const ids = session.ctx.ids.scope(`${key}/`);
  // A story's tracked changes stay as E1 projected them (ED-005 sections 5.2 and 10): preserved, not shown.
  const ctx = newContext(session.schema, ids, 'hidden');
  let blocks = projectBlocks(content ?? [], ctx, key);
  ids.seal();
  if (blocks.length === 0) blocks = [session.schema.nodes['paragraph']!.create()];
  return { key, doc: session.schema.topNodeType.create(null, blocks), ctx };
}

/**
 * A comment's first paragraph's `w14:paraId` put on the element its story projected from, so a
 * later save, which rebuilds the paragraph from that element, keeps the id the side parts are
 * keyed by (ED-005 section 7.4).
 */
function keepParaId(record: CommentRecord): void {
  const first = (record.element.content ?? []).find((el) => typeNameOf(el) === 'org_docx4j_wml.P')?.value as wml.P | undefined;
  const sourceId = record.doc.firstChild?.attrs['sourceId'] as string | undefined;
  const original = sourceId ? record.ctx.ids.get(sourceId)?.value as wml.P | undefined : undefined;
  if (first?.paraId && original && original !== first) original.paraId = first.paraId;
}

/** The ids the document references, per kind, in order of first reference. */
function referencedIds(doc: PMNode, kind: NoteKind): number[] {
  const out: number[] = [];
  const seen = new Set<number>();
  doc.descendants((node) => {
    if (node.type.name !== 'footnoteRef' || node.attrs['kind'] !== kind) return !node.isLeaf;
    const id = Number(node.attrs['footnoteId']);
    if (!seen.has(id)) { seen.add(id); out.push(id); }
    return false;
  });
  return out;
}

export class Stories {
  private readonly notesCache = new Map<NoteKind, Promise<NotesInfo | undefined>>();
  private readonly notesLoaded = new Map<NoteKind, NotesInfo>();
  private commentsCache: Promise<CommentsInfo> | undefined;
  private commentsLoaded: CommentsInfo | undefined;
  private headersCache: Promise<HeaderFooter[]> | undefined;
  private headersLoaded: HeaderFooter[] | undefined;
  private settingsCache: Promise<wml.CTSettings | undefined> | undefined;
  private readonly listeners = new Set<() => void>();

  constructor(private readonly session: EditorDocument) {}

  /** Footnotes in the order their references appeared at open; ids -1 and 0 never; orphans last; notes made here after. */
  async footnotes(): Promise<Note[]> {
    const info = await this.notesInfo('footnote');
    if (info) await this.adoptNotes('footnote', info);
    return info?.notes ?? [];
  }

  async endnotes(): Promise<Note[]> {
    const info = await this.notesInfo('endnote');
    if (info) await this.adoptNotes('endnote', info);
    return info?.notes ?? [];
  }

  /**
   * The notes another session made that this one does not know: a fragment in the map with no
   * note of that id here, whose element is in `sources` (written by `newNote`). Each is made a
   * note of this session with the sender's ids, as it would have been had it been made here: a
   * peer's note in a room, or a recovered session's own unsaved one (ED-004 section 3.7).
   */
  private async adoptNotes(kind: NoteKind, info: NotesInfo): Promise<void> {
    const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
    for (const key of map.keys()) {
      const id = Number(key);
      if (!Number.isFinite(id) || info.byId.has(id)) continue;
      const entry = this.session.store.sources.get(noteSourceKey(kind, id)) as SourceEntry | undefined;
      if (!entry) continue;
      const parsed = (await unmarshalString(entry.xml)).value as wml.CTFootnotes & wml.CTEndnotes;
      const element = (kind === 'footnote' ? parsed.footnote : parsed.endnote)?.[0];
      if (!element || info.byId.has(id)) continue;
      assignPathIds(element as unknown as Element, entry.ids, this.session.ctx.ids);
      linkParents(element, undefined);
      const note: Note = { id, kind, orphan: false, element, isNew: true, ...projectStory(this.session, element.content as Element[], `${kind}:${id}`) };
      info.notes.push(note);
      info.byId.set(id, note);
    }
  }

  /** The notes of a kind, the part read on the first call and kept (a save writes into its tree). */
  private notesInfo(kind: NoteKind): Promise<NotesInfo | undefined> {
    let cached = this.notesCache.get(kind);
    if (!cached) {
      cached = this.readNotes(kind).then((info) => { if (info) { this.notesLoaded.set(kind, info); this.notify(); } return info; });
      this.notesCache.set(kind, cached);
    }
    return cached;
  }

  private async readNotes(kind: NoteKind): Promise<NotesInfo | undefined> {
    const main = this.session.main;
    const part = kind === 'footnote' ? main.footnotesPart : main.endnotesPart;
    if (!part) return undefined;
    let tree: wml.CTFootnotes | wml.CTEndnotes;
    try {
      tree = await part.readContents();
    } catch {
      return undefined;
    }
    const items = (kind === 'footnote' ? (tree as wml.CTFootnotes).footnote : (tree as wml.CTEndnotes).endnote) ?? [];
    const elements = new Map<number, wml.CTFtnEdn>();
    for (const n of items) if (n.id > 0) elements.set(n.id, n);
    const notes: Note[] = [];
    const byId = new Map<number, Note>();
    const make = (id: number, element: wml.CTFtnEdn, orphan: boolean): void => {
      const note: Note = { id, kind, orphan, element, isNew: false, ...projectStory(this.session, element.content as Element[] | undefined, `${kind}:${id}`) };
      notes.push(note);
      byId.set(id, note);
    };
    // The baseline decides what "orphan" means: a reference deleted before the panel opened still counts as deleted.
    for (const id of referencedIds(this.session.imported, kind)) {
      const el = elements.get(id);
      if (el && !byId.has(id)) make(id, el, false);
    }
    for (const [id, el] of elements) if (!byId.has(id)) make(id, el, true);
    return { part, tree, notes, byId };
  }

  /**
   * The notes a panel lists for the document as it is now: the referenced ones in reference
   * order (a note the body no longer references is left out, since a save removes it), then
   * the orphans, which a save keeps. Synchronous, over the parts already read: call after
   * `footnotes()` or `endnotes()` resolved, or get an empty list.
   */
  orderedNotes(kind: NoteKind, doc: PMNode): Note[] {
    const info = this.notesLoaded.get(kind);
    if (!info) return [];
    const out: Note[] = [];
    for (const id of referencedIds(doc, kind)) { const n = info.byId.get(id); if (n) out.push(n); }
    for (const n of info.notes) if (n.orphan && !out.includes(n)) out.push(n);
    return out;
  }

  /**
   * The stories find searches after the body, in Word's order (ED-005 section 3.7): the footnotes
   * and then the endnotes as `orderedNotes` lists them for `doc`, then the headers and footers.
   * Reads the parts on first use; comments are not searched, as Word's find does not search them.
   */
  async inFindOrder(doc: PMNode = this.session.doc): Promise<Story[]> {
    await Promise.all([this.footnotes(), this.endnotes()]);
    return [...this.orderedNotes('footnote', doc), ...this.orderedNotes('endnote', doc), ...(await this.headers())];
  }

  /** The `sourceId` of the first reference to a note in `doc`, when there is one. */
  referenceOf(doc: PMNode, kind: NoteKind, id: number): string | undefined {
    let found: string | undefined;
    doc.descendants((node) => {
      if (found) return false;
      if (node.type.name === 'footnoteRef' && node.attrs['kind'] === kind && Number(node.attrs['footnoteId']) === id) found = String(node.attrs['sourceId']);
      return !found && !node.isLeaf;
    });
    return found;
  }

  /** A note by kind and id, once the part is read. */
  noteOf(kind: NoteKind, id: number): Note | undefined {
    return this.notesLoaded.get(kind)?.byId.get(id);
  }

  /** The session made the notes part (`EditorDocument.ensureNotesPart`): the next read sees it. */
  partAdded(kind: NoteKind): void {
    this.notesCache.delete(kind);
    this.notesLoaded.delete(kind);
  }

  /** The next id: above every id in the part, in the map and in the body's references. */
  private nextNoteId(kind: NoteKind, info: NotesInfo): number {
    // In a room a random id, so two peers inserting a note at once never share one (ED-004 section 3.4).
    if (this.session.inRoom) {
      const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
      const used = new Set<number>([...info.byId.keys(), ...info.notes.map((n) => n.id), ...[...map.keys()].map(Number), ...referencedIds(this.session.doc, kind)]);
      return randomId(used);
    }
    let max = 0;
    for (const id of info.byId.keys()) max = Math.max(max, id);
    const list = (kind === 'footnote' ? (info.tree as wml.CTFootnotes).footnote : (info.tree as wml.CTEndnotes).endnote) ?? [];
    for (const n of list) max = Math.max(max, n.id);
    const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
    for (const key of map.keys()) max = Math.max(max, Number(key) || 0);
    for (const id of referencedIds(this.session.doc, kind)) max = Math.max(max, id);
    return max + 1;
  }

  /**
   * A new note (ED-003 section 6.3): the next id, one paragraph in the note's text style holding
   * the note's own mark in the reference style and the text after a space, as Word writes one,
   * built with the builders as a `w:footnote` element, projected like any note and loaded into
   * its fragment. It joins the part at the first save that finds the body referencing it (6.2);
   * until then it is content in the map.
   */
  async newNote(kind: NoteKind, text = ''): Promise<Note> {
    const info = await this.notesInfo(kind);
    if (!info) throw new Error(`the document has no ${kind}s part`);
    const id = this.nextNoteId(kind, info);
    const styles = NOTE_STYLES[kind];
    const mark = runOf([(kind === 'footnote' ? el.footnoteRef({}) : el.endnoteRef({})) as Element], { rStyle: { val: styles.reference } });
    const paragraph = p([mark, runOf([t(` ${text}`)])], { style: styles.text });
    const element: wml.CTFtnEdn = { TYPE_NAME: 'org_docx4j_wml.CTFtnEdn', id, content: [paragraph] as wml.CTFtnEdn['content'] };
    linkParents(element, undefined);
    const note: Note = { id, kind, orphan: false, element, isNew: true, ...projectStory(this.session, element.content as Element[], `${kind}:${id}`) };
    info.notes.push(note);
    info.byId.set(id, note);
    // The note's element into `sources` with its fragment, in one transaction: another session
    // (a peer, a recovery) adopts the note from it rather than dropping its reference.
    const wrapper = kind === 'footnote' ? el.footnotes({ footnote: [element] }) : el.endnotes({ endnote: [element] });
    const entry: SourceEntry = { xml: await marshalString(wrapper as unknown as Element), ids: pathIds(element as unknown as Element, this.session.ctx.ids) };
    this.session.store.ydoc.transact(() => {
      this.session.store.sources.set(noteSourceKey(kind, id), entry);
      this.fragmentOf(note, true);
    }, IMPORT_ORIGIN);
    this.notify();
    return note;
  }

  // ---------------------------------------------------------------- fragments (ED-003 section 6.2)

  private mapOf(story: Story): { map: Y.Map<Y.XmlFragment>; key: string } {
    const store = this.session.store;
    const colon = story.key.indexOf(':');
    const kind = story.key.slice(0, colon);
    const rest = story.key.slice(colon + 1);
    switch (kind) {
      case 'footnote': return { map: store.footnotes, key: rest };
      case 'endnote': return { map: store.endnotes, key: rest };
      case 'header':
      case 'footer': return { map: store.headers, key: rest };
      case 'comment': return { map: store.comments, key: rest };
      default: throw new Error(`${story.key} has no fragment`);
    }
  }

  /**
   * The story's fragment: the one the side map holds (a recovered session, another view over
   * the same header), else a new one loaded with the story's projection under the import
   * origin, so binding a view is neither an undo step nor a dirty mark.
   */
  /**
   * The story's fragment, made on first binding outside a room (ED-003 section 6.2). In a room it
   * must be there already: the creator made every note's and header's at Share and a command
   * makes a new note's with the note, so a missing one is a `RoomNotReadyError`, never a new
   * fragment that another peer's would overwrite (ED-004 section 3.9). `create` is the Share's
   * and the new note's.
   */
  fragmentOf(story: Story, create = !this.session.inRoom): Y.XmlFragment {
    const { map, key } = this.mapOf(story);
    return fragmentIn(this.session.store, map, key, () => story.doc, create);
  }

  /** Every note's, header's and comment's fragment made, as Share does before the room is shared. */
  async fillAll(): Promise<number> {
    const comments = [...(await this.commentsInfo()).records.values()];
    const stories: Story[] = [...(await this.footnotes()), ...(await this.endnotes()), ...(await this.headers()), ...comments];
    for (const s of stories) this.fragmentOf(s, true);
    return stories.length;
  }

  /** Whether an editable view has bound the story (or a recovered session carried it). */
  hasFragment(story: Story): boolean {
    const { map, key } = this.mapOf(story);
    return map.has(key);
  }

  /** The story as it is now: its fragment's document when it has one, else its projection. */
  currentDoc(story: Story): PMNode {
    const { map, key } = this.mapOf(story);
    const fragment = map.get(key);
    return fragment ? readFragment(fragment, this.session.schema) : story.doc;
  }

  /** Replace a story's document as one local transaction (headless editing; the view binds through the fragment). */
  replace(story: Story, doc: PMNode, origin: unknown = LOCAL_ORIGIN): void {
    replaceFragment(this.session.store, this.fragmentOf(story), doc, origin);
  }

  /** The element one block of a story stands for now (reveal codes over a note or a header). Pure. */
  exportBlock(story: Story, node: PMNode): Element {
    const doc = this.session.schema.topNodeType.create(null, [node]);
    return exportBlocks(doc, story.ctx, { content: [] }, this.session.controlEntry).content[0]!;
  }

  /** Every note and header read so far: what the stylesheet covers beside the body (ED-003 section 6.4). */
  loadedStories(): Story[] {
    const out: Story[] = [];
    for (const info of this.notesLoaded.values()) out.push(...info.notes);
    if (this.headersLoaded) out.push(...this.headersLoaded);
    return out;
  }

  /** Something a panel lists changed outside a document transaction (a part read, a note made); `notify` is the model's side. */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  protected notify(): void {
    for (const fn of this.listeners) fn();
  }

  // ---------------------------------------------------------------- the save (ED-003 section 6.2)

  /**
   * Write every story part with a change into its part and answer the part names written: a
   * note whose fragment differs from its baseline, a `w:footnote` referenced at open and not
   * in `doc`, a note made here that `doc` references, a header or footer whose fragment
   * differs. A part with no change keeps its bytes.
   */
  async exportAll(doc: PMNode, ids?: RevisionIds): Promise<string[]> {
    const touched: string[] = [];
    for (const kind of ['footnote', 'endnote'] as const) {
      const name = await this.exportNotes(kind, doc, ids);
      if (name) touched.push(name);
    }
    touched.push(...await this.exportHeaders(ids));
    return touched;
  }

  private async exportNotes(kind: NoteKind, doc: PMNode, ids?: RevisionIds): Promise<string | undefined> {
    const info = await this.notesInfo(kind);
    if (!info) return undefined;
    await this.adoptNotes(kind, info);
    const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
    const referenced = new Set(referencedIds(doc, kind));
    const list = (kind === 'footnote' ? (info.tree as wml.CTFootnotes).footnote : (info.tree as wml.CTEndnotes).endnote) ?? [];
    let changed = false;
    const out: wml.CTFtnEdn[] = [];
    for (const el of list) {
      const note = el.id > 0 ? info.byId.get(el.id) : undefined;
      if (note && !note.isNew && !note.orphan && !referenced.has(el.id)) { changed = true; continue; }   // Word's rule: the reference gone, the note goes
      if (note && this.exportStoryContent(note, map.get(String(note.id)), ids)) changed = true;
      out.push(el);
    }
    // The notes made in sessions join the part in order of id, so every peer writes the same part (ED-004 section 3.3).
    for (const note of [...info.notes].sort((a, b) => a.id - b.id)) {
      if (!note.isNew || !referenced.has(note.id) || out.includes(note.element)) continue;
      this.exportStoryContent(note, map.get(String(note.id)), ids);
      out.push(note.element);
      changed = true;
    }
    if (!changed) return undefined;
    if (kind === 'footnote') (info.tree as wml.CTFootnotes).footnote = out; else (info.tree as wml.CTEndnotes).endnote = out;
    linkParents(info.tree, undefined);
    (info.part as { setContents(value: unknown): void }).setContents(info.tree);
    for (const note of info.notes) if (out.includes(note.element)) note.isNew = false;
    return info.part.partName.toString();
  }

  /**
   * The ids a save writes for the notes made in a room (ED-004 section 10.8 item 6). In a room a
   * new note takes a random id in [2^20, 2^31 - 1], which keeps two peers' notes apart (section
   * 3.4) but which Word 2010 refuses in a notes part. So the file gets each such note, in the
   * order the document references them, the smallest positive id no other note in the part has;
   * the notes the base file brought keep theirs. The same document gives every peer the same ids.
   * Called after `exportAll`, over the parts as they will be written.
   */
  noteIdsForFile(doc: PMNode): NoteIds {
    const out: NoteIds = new Map();
    for (const kind of ['footnote', 'endnote'] as const) {
      const info = this.notesLoaded.get(kind);
      if (!info) continue;
      const list = ((kind === 'footnote' ? (info.tree as wml.CTFootnotes).footnote : (info.tree as wml.CTEndnotes).endnote) ?? []) as { id: number }[];
      const madeHere = (id: number): boolean => this.session.store.sources.has(noteSourceKey(kind, id));
      const inPart = new Set(list.map((el) => Number(el.id)));
      const used = new Set([...inPart].filter((id) => !madeHere(id)));
      const map = new Map<number, number>();
      let next = 1;
      for (const id of referencedIds(doc, kind)) {
        if (!madeHere(id) || !inPart.has(id)) continue;
        while (used.has(next)) next++;
        map.set(id, next);
        used.add(next);
      }
      if (map.size > 0) out.set(kind, map);
    }
    return out;
  }

  /** Write `ids` into the notes parts' trees for a save; what it returns puts the room's ids back. */
  applyNoteIds(ids: NoteIds): () => void {
    const undo: (() => void)[] = [];
    for (const [kind, map] of ids) {
      const info = this.notesLoaded.get(kind);
      if (!info) continue;
      const list = ((kind === 'footnote' ? (info.tree as wml.CTFootnotes).footnote : (info.tree as wml.CTEndnotes).endnote) ?? []) as { id: number }[];
      const order = [...list];
      const slots: number[] = [];
      for (let i = 0; i < list.length; i++) {
        const el = list[i]!;
        const to = map.get(Number(el.id));
        if (to === undefined) continue;
        const from = el.id;
        el.id = to;
        slots.push(i);
        undo.push(() => { el.id = from; });
      }
      // The renumbered notes in the places they hold, in order of their new ids, as Word writes a
      // notes part (they were appended in order of their random ids).
      const renumbered = slots.map((i) => list[i]!).sort((a, b) => a.id - b.id);
      slots.forEach((slot, k) => { list[slot] = renumbered[k]!; });
      undo.push(() => { order.forEach((el, i) => { list[i] = el; }); });
    }
    return () => { for (const f of [...undo].reverse()) f(); };
  }

  /**
   * A notes, header or footer part's tree for Select by XPath (ED-005 section 3.5): every bound
   * story's fragment exported into its element as a save does, so the expression sees the text on
   * screen; the part's root as an element, and each story's content with its document. Undefined
   * for a part that is not one of these or that the document lacks.
   */
  async selectTree(partName: string): Promise<{ root: Element; containers: { story: Story; content: Element[]; doc: PMNode }[] } | undefined> {
    for (const kind of ['footnote', 'endnote'] as const) {
      const info = await this.notesInfo(kind);
      if (!info || info.part.partName.toString() !== partName) continue;
      await this.adoptNotes(kind, info);
      const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
      const list = (kind === 'footnote' ? (info.tree as wml.CTFootnotes).footnote : (info.tree as wml.CTEndnotes).endnote) ?? [];
      const containers: { story: Story; content: Element[]; doc: PMNode }[] = [];
      for (const element of list) {
        const note = element.id > 0 ? info.byId.get(element.id) : undefined;
        if (!note || note.element !== element) continue;
        this.exportStoryContent(note, map.get(String(note.id)));
        containers.push({ story: note, content: (element.content ?? []) as Element[], doc: this.currentDoc(note) });
      }
      linkParents(info.tree, undefined);
      const root = (kind === 'footnote' ? el.footnotes(info.tree as wml.CTFootnotes) : el.endnotes(info.tree as wml.CTEndnotes)) as unknown as Element;
      return { root, containers };
    }
    const hf = (await this.headers()).find((h) => h.part.partName.toString() === partName);
    if (!hf) return undefined;
    this.exportStoryContent(hf, this.session.store.headers.get(partName));
    linkParents(hf.element, undefined);
    const root = (hf.kind === 'header' ? el.hdr(hf.element as wml.Hdr) : el.ftr(hf.element as wml.Ftr)) as unknown as Element;
    return { root, containers: [{ story: hf, content: (hf.element.content ?? []) as Element[], doc: this.currentDoc(hf) }] };
  }

  /** The notes, header and footer parts Select by XPath can run over, by part name, in Word's order. */
  async selectableParts(): Promise<string[]> {
    const out: string[] = [];
    for (const part of [this.session.main.footnotesPart, this.session.main.endnotesPart]) if (part) out.push(part.partName.toString());
    for (const hf of await this.headers()) out.push(hf.part.partName.toString());
    return out;
  }

  /**
   * Rebuild a story's element content from its fragment when the fragment differs from the
   * baseline; whether it did. `ids` is a save's id source, shared with the body's export (ED-005
   * section 5.7); without one the story counts on from its own highest id.
   */
  private exportStoryContent(story: Story & { element: { content?: unknown } }, fragment: Y.XmlFragment | undefined, ids?: RevisionIds): boolean {
    if (!fragment) return false;
    const current = readFragment(fragment, this.session.schema);
    if (current.eq(story.doc)) return false;
    const result = exportBlocks(current, story.ctx, story.element as { content?: Element[] }, this.session.controlEntry, ids);
    if (!result.changed) return false;
    story.element.content = result.content;
    return true;
  }

  private async exportHeaders(ids?: RevisionIds): Promise<string[]> {
    const map = this.session.store.headers;
    if (map.size === 0) return [];
    const all = await this.headers();
    const touched: string[] = [];
    for (const [partName, fragment] of map) {
      const hf = all.find((h) => h.part.partName.toString() === partName);
      if (!hf || !this.exportStoryContent(hf, fragment, ids)) continue;
      linkParents(hf.element, undefined);
      (hf.part as { setContents(value: unknown): void }).setContents(hf.element);
      touched.push(partName);
    }
    return touched;
  }

  /**
   * The self-check over the stories (ED-001 section 11 item 1): every bound note the document
   * keeps, and every bound header, re-projected from the saved bytes and compared with the
   * live fragment; the first difference names its story.
   */
  async selfCheck(reloaded: EditorDocument, doc: PMNode, ids: NoteIds = new Map(), commentIds: CommentIds = new Map()): Promise<SelfCheckResult> {
    const comments = await this.selfCheckComments(reloaded, commentIds);
    if (!comments.equal) return comments;
    for (const kind of ['footnote', 'endnote'] as const) {
      const info = this.notesLoaded.get(kind);
      const map = kind === 'footnote' ? this.session.store.footnotes : this.session.store.endnotes;
      if (!info || map.size === 0) continue;
      // `doc` carries the ids the file was written with; a note's id in the file is `inFile`.
      const inFile = (id: number): number => ids.get(kind)?.get(id) ?? id;
      const referenced = new Set(referencedIds(doc, kind));
      const theirs = kind === 'footnote' ? await reloaded.stories.footnotes() : await reloaded.stories.endnotes();
      for (const key of map.keys()) {
        const note = info.byId.get(Number(key));
        if (!note || (!note.orphan && !referenced.has(inFile(note.id)))) continue;
        const other = theirs.find((n) => n.id === inFile(note.id));
        if (!other) return { equal: false, difference: { index: -1, live: 'present', reloaded: 'missing', story: note.key } };
        const result = compareProjections(this.currentDoc(note), other.doc);
        if (!result.equal) return { equal: false, difference: { ...result.difference!, story: note.key } };
      }
    }
    if (this.session.store.headers.size > 0) {
      const mine = await this.headers();
      const theirs = await reloaded.stories.headers();
      for (const partName of this.session.store.headers.keys()) {
        const hf = mine.find((h) => h.part.partName.toString() === partName);
        if (!hf) continue;
        const other = theirs.find((h) => h.part.partName.toString() === partName);
        if (!other) return { equal: false, difference: { index: -1, live: 'present', reloaded: 'missing', story: hf.key } };
        const result = compareProjections(this.currentDoc(hf), other.doc);
        if (!result.equal) return { equal: false, difference: { ...result.difference!, story: hf.key } };
      }
    }
    return { equal: true };
  }

  // ---------------------------------------------------------------- comments (ED-005 section 7)

  /**
   * The comments as a panel lists them now (`commentThreads`), the part read on the first call:
   * top-level ones in id order, replies under their parent, orphans last.
   */
  async comments(doc?: PMNode): Promise<CommentThread[]> {
    await this.commentsInfo();
    return this.commentThreads(doc ?? this.session.doc);
  }

  /** The comments part read on the first call and kept (a save writes into its tree), a peer's new comments adopted. */
  private async commentsInfo(): Promise<CommentsInfo> {
    const info = await (this.commentsCache ??= this.readComments().then((i) => { this.commentsLoaded = i; this.notify(); return i; }));
    if (await this.adoptComments(info)) this.notify();
    return info;
  }

  /** A comment by id, once the part is read (or made in this session). */
  commentOf(id: string): CommentRecord | undefined {
    return this.commentsLoaded?.records.get(id);
  }

  /**
   * The comments for `doc`, synchronous over what is read: each record with its
   * `commentThreads` entry laid over; left out, one marked deleted, and one `doc` no longer
   * anchors unless it was an orphan at open (a save drops both, section 7.4, as Word does when
   * a comment's text is deleted); replies under a parent listed, and with it. Empty before the
   * first `comments()` resolves.
   */
  commentThreads(doc: PMNode): CommentThread[] {
    const info = this.commentsLoaded;
    if (!info) return [];
    const anchors = commentAnchors(doc, this.session.ctx.ids);
    const all = new Map<string, CommentThread>();
    for (const record of info.records.values()) {
      const entry = readThreadEntry(this.session.store, record.id);
      if (entry?.deleted) continue;
      if (!record.orphan && !anchors.has(record.id)) continue;
      all.set(record.id, {
        key: record.key, doc: record.doc, ctx: record.ctx, record,
        id: record.id,
        author: entry?.author ?? record.base.author,
        initials: entry?.initials ?? record.base.initials,
        date: entry?.date !== undefined ? new Date(entry.date) : record.base.date,
        parentId: entry?.parent !== undefined ? entry.parent : record.base.parentId,
        done: entry?.done ?? record.base.done,
        orphan: record.orphan, isNew: record.isNew, replies: [],
      });
    }
    const top: CommentThread[] = [];
    const byId = (a: CommentThread, b: CommentThread): number => Number(a.id) - Number(b.id);
    const sorted = [...all.values()].sort(byId);
    for (const t of sorted) {
      const parent = t.parentId !== null ? all.get(t.parentId) : undefined;
      if (parent && parent !== t) parent.replies.push(t);
      else if (t.parentId === null || !info.records.has(t.parentId)) top.push(t);
      // A reply whose parent is gone goes with it, as the engine's delete takes a thread's replies.
    }
    return top.sort((a, b) => Number(a.orphan) - Number(b.orphan) || byId(a, b));
  }

  private async readComments(): Promise<CommentsInfo> {
    const main = this.session.main;
    if (!main.commentsPart) return { tree: undefined, records: new Map() };
    let comments: wml.Comments;
    try {
      // The part's own tree when a script has unmarshalled it, else a private one (decision 5).
      comments = await main.commentsPart.readContents();
    } catch {
      return { tree: undefined, records: new Map() };
    }
    return { tree: comments, records: await this.recordsFrom(comments) };
  }

  /** A record per comment of a comments part's tree, threaded through `commentsExtended`, each projected into its own story. */
  private async recordsFrom(comments: wml.Comments): Promise<Map<string, CommentRecord>> {
    const main = this.session.main;
    const records = new Map<string, CommentRecord>();
    let extended: w15.CTCommentsEx | undefined;
    if (main.commentsExtendedPart) {
      try { extended = await main.commentsExtendedPart.readContents(); } catch { extended = undefined; }
    }
    const exByParaId = new Map<string, w15.CTCommentEx>();
    for (const ex of extended?.commentEx ?? []) exByParaId.set(ex.paraId, ex);
    const firstParaId = (c: wml.Comments.Comment): string | undefined => {
      for (const el of c.content ?? []) if (typeNameOf(el) === 'org_docx4j_wml.P') return (el.value as wml.P).paraId;
      return undefined;
    };
    const paraIdToCommentId = new Map<string, string>();
    for (const c of comments.comment ?? []) { const pid = firstParaId(c); if (pid) paraIdToCommentId.set(pid.toUpperCase(), String(c.id)); }
    const datesUtc = await commentDatesUtc(main);
    // An anchor at open is a range start or a reference run: Word writes the reference alone for an empty range.
    const anchored = commentAnchors(this.session.imported, this.session.ctx.ids);
    for (const c of comments.comment ?? []) {
      const id = String(c.id);
      const pid = firstParaId(c);
      const ex = pid ? exByParaId.get(pid) ?? exByParaId.get(pid.toUpperCase()) ?? exByParaId.get(pid.toLowerCase()) : undefined;
      const parentId = ex?.paraIdParent ? paraIdToCommentId.get(ex.paraIdParent.toUpperCase()) ?? null : null;
      records.set(id, {
        id, element: c, isNew: false,
        orphan: !anchored.has(id) && !this.session.ctx.commentIds.has(id),
        base: { author: c.author ?? '', initials: c.initials ?? '', date: commentDate(c, datesUtc), parentId, done: ex?.done === '1' || ex?.done === 'true' },
        ...projectStory(this.session, c.content as Element[] | undefined, `comment:${id}`),
      });
    }
    return records;
  }

  /** What each comment was when a tree run's callback began: its element and its content array, which a script replaces to change the text. */
  private runSnapshot: Map<string, { element: wml.Comments.Comment; content: unknown }> | undefined;

  /**
   * Before a tree run's callback (ED-005 section 7.4): the session's comment changes written into
   * the comments part as a save writes them, so a script sees the comments on screen, and what
   * each comment is, so that `afterTreeRun` sees what the script changed.
   */
  async beforeTreeRun(doc: PMNode, body: wml.Body, ids?: RevisionIds): Promise<void> {
    await this.exportComments(doc, body, ids);
    const info = this.commentsLoaded;
    this.runSnapshot = info ? new Map([...info.records.values()].map((r) => [r.id, { element: r.element, content: r.element.content }])) : undefined;
  }

  /**
   * After a tree run's callback: when the script read the comments part, the records made again
   * from it as it is (a comment the script made listed, one it deleted gone), a comment the script
   * left alone keeping its projection (its ids, so its paragraph ids survive the next export), one
   * whose text it changed loaded into its fragment, and a thread's resolved state the part's, an
   * entry that says otherwise set to match, under the run's origin. A comment the script made is
   * published as `newComment` publishes one (its element in `sources`, its fragment, its thread
   * entry, under the import origin), so that another session in the room adopts it: before E4.f
   * step F5 found it (ED-005 section 12.29), a peer never saw a comment a script or the agent made.
   */
  async afterTreeRun(origin: unknown): Promise<void> {
    const snapshot = this.runSnapshot;
    this.runSnapshot = undefined;
    const part = this.session.main.commentsPart;
    const info = this.commentsLoaded;
    if (!part || !info || !part.isUnmarshalled) return;
    const live = await part.getContents();
    const fresh = await this.recordsFrom(live);
    const store = this.session.store;
    // The comments the run made: in no snapshot, no record and no thread entry before it.
    const made = [...fresh.values()].filter((r) => !snapshot?.has(r.id) && !info.records.has(r.id) && !store.commentThreads.has(r.id));
    const entries = new Map<string, SourceEntry>();
    for (const record of made) {
      // A copy in the entry's wrapper, so the part's own element keeps its parent; the paths are the same.
      const wrapper = el.comments({ comment: [deepCopy(record.element)] });
      entries.set(record.id, { xml: await marshalString(wrapper as unknown as Element), ids: pathIds(record.element as unknown as Element, this.session.ctx.ids) });
    }
    for (const [id, record] of fresh) {
      const before = snapshot?.get(id);
      const old = info.records.get(id);
      if (old && before && record.element === before.element && record.element.content === before.content) {
        fresh.set(id, { ...old, element: record.element, base: record.base, orphan: record.orphan, isNew: false });
      } else if (store.comments.has(id)) {
        replaceFragment(store, store.comments.get(id)!, record.doc, origin);
      }
    }
    store.ydoc.transact(() => {
      for (const record of fresh.values()) {
        const entry = readThreadEntry(store, record.id);
        if (entry?.done !== undefined && entry.done !== record.base.done) writeThreadFields(store, record.id, { done: record.base.done });
      }
    }, origin);
    if (made.length > 0) {
      store.ydoc.transact(() => {
        for (const record of made) {
          store.sources.set(commentSourceKey(record.id), entries.get(record.id)!);
          this.fragmentOf(record, true);
          const { author, initials, date, parentId, done } = record.base;
          writeThreadFields(store, record.id, { author, initials, date: (date ?? new Date()).toISOString(), parent: parentId, done });
        }
      }, IMPORT_ORIGIN);
    }
    info.tree = live;
    info.records = fresh;
    this.notify();
  }

  /**
   * A comment made in this session (ED-005 section 7.2): the next id outside a room (above every
   * comment and every anchor, the engine's `nextCommentId` rule), a random one in a room; its
   * element built as Word writes one, projected, and loaded into its fragment; its thread entry
   * and its element in `sources` in the same transaction, under the import origin, so another
   * session adopts it and undo leaves it (an undone insert takes its anchor, and the listing and
   * the save leave out a new comment with none).
   */
  async newComment(made: NewComment): Promise<CommentRecord> {
    const info = await this.commentsInfo();
    const id = String(this.nextCommentId(info));
    const date = made.date ?? new Date();
    const initials = made.initials ?? initialsOf(made.author);
    const element = commentElement(Number(id), made.author, initials, date, made.text ?? '');
    const record: CommentRecord = {
      id, element, isNew: true, orphan: false,
      base: { author: made.author, initials, date, parentId: made.parent ?? null, done: false },
      ...projectStory(this.session, element.content as Element[], `comment:${id}`),
    };
    info.records.set(id, record);
    const wrapper = el.comments({ comment: [element] });
    const entry: SourceEntry = { xml: await marshalString(wrapper as unknown as Element), ids: pathIds(element as unknown as Element, this.session.ctx.ids) };
    this.session.store.ydoc.transact(() => {
      this.session.store.sources.set(commentSourceKey(id), entry);
      this.fragmentOf(record, true);
      writeThreadFields(this.session.store, id, { author: made.author, initials, date: date.toISOString(), parent: made.parent ?? null, done: false });
    }, IMPORT_ORIGIN);
    this.notify();
    return record;
  }

  /** The next comment id: above every comment in the part, in the maps and anchored in the body; random in a room (ED-004 section 3.4). */
  private nextCommentId(info: CommentsInfo): number {
    const store = this.session.store;
    const used = new Set<number>();
    for (const c of info.tree?.comment ?? []) if (typeof c.id === 'number') used.add(c.id);
    for (const id of info.records.keys()) used.add(Number(id));
    for (const key of [...store.comments.keys(), ...store.commentThreads.keys()]) used.add(Number(key));
    for (const id of commentAnchors(this.session.doc, this.session.ctx.ids)) used.add(Number(id));
    for (const id of this.session.ctx.commentIds) used.add(Number(id));
    const known = [...used].filter((n) => Number.isFinite(n));
    if (this.session.inRoom) return randomId(new Set(known));
    return known.length === 0 ? 0 : Math.max(...known) + 1;
  }

  /**
   * The comments another session made that this one does not know: an entry in
   * `commentThreads` with no record here, whose element is in `sources` (written by
   * `newComment`), made a record with the maker's ids, as `adoptNotes` does. Whether any was.
   */
  private async adoptComments(info: CommentsInfo): Promise<boolean> {
    let adopted = false;
    for (const id of this.session.store.commentThreads.keys()) {
      if (info.records.has(id)) continue;
      const entry = this.session.store.sources.get(commentSourceKey(id)) as SourceEntry | undefined;
      if (!entry) continue;
      const parsed = (await unmarshalString(entry.xml)).value as wml.Comments;
      const element = parsed.comment?.[0];
      if (!element || info.records.has(id)) continue;
      assignPathIds(element as unknown as Element, entry.ids, this.session.ctx.ids);
      linkParents(element, undefined);
      const fields = readThreadEntry(this.session.store, id) ?? {};
      info.records.set(id, {
        id, element, isNew: true, orphan: false,
        base: { author: fields.author ?? element.author ?? '', initials: fields.initials ?? element.initials ?? '', date: fields.date ? new Date(fields.date) : commentDate(element, new Map()), parentId: fields.parent ?? null, done: false },
        ...projectStory(this.session, element.content as Element[], `comment:${id}`),
      });
      adopted = true;
    }
    return adopted;
  }

  // ---------------------------------------------------------------- comments at save (ED-005 section 7.4)

  /** The comments the last save wrote or changed, for its self-check. */
  private commentsWritten = new Set<string>();

  /**
   * The comments part and its side parts written for a save (ED-005 section 7.4), after the
   * body's export has put its markers in `body`: nothing when no comment changed. The session's
   * tree goes to the comments part (as a notes part's does), made when the document has none;
   * a file comment's text from its fragment, the first paragraph keeping its `w14:paraId`; one
   * dropped (deleted, or no longer anchored unless an orphan at open) through the engine's
   * `Comment.delete`, which takes its replies, markers and side entries; a comment made here
   * through the engine's `addCommentEntry` (core-ts 0.3.0, CR-002 section 36), parents before
   * replies, its story's element then the engine's; `done` through the engine's `resolved`. The
   * engine's views reach the parts: `Body.getComments()` gives one for an anchored comment, and
   * `addCommentEntry` another. The part names written, and whether the body's tree changed.
   */
  async exportComments(doc: PMNode, body: wml.Body, ids?: RevisionIds): Promise<{ touched: string[]; bodyChanged: boolean }> {
    this.commentsWritten = new Set();
    const store = this.session.store;
    if (!this.commentsLoaded && store.commentThreads.size === 0 && store.comments.size === 0) return { touched: [], bodyChanged: false };
    const info = await this.commentsInfo();
    const anchors = commentAnchors(doc, this.session.ctx.ids);
    const entryOf = (r: CommentRecord) => readThreadEntry(store, r.id);
    const parentOf = (r: CommentRecord): string | null => { const e = entryOf(r); return e?.parent !== undefined ? e.parent : r.base.parentId; };
    const kept = (r: CommentRecord): boolean => !entryOf(r)?.deleted && (r.orphan || anchors.has(r.id));
    const made: CommentRecord[] = [];
    const dropped: CommentRecord[] = [];
    const edited: CommentRecord[] = [];
    const resolved: [CommentRecord, boolean][] = [];
    for (const r of info.records.values()) {
      if (r.isNew) {
        const parent = parentOf(r);
        const parentRecord = parent === null ? undefined : info.records.get(parent);
        if (kept(r) && (parent === null || (parentRecord !== undefined && kept(parentRecord)))) made.push(r);
        continue;
      }
      if (!kept(r)) { dropped.push(r); continue; }
      if (this.exportStoryContent(r, store.comments.get(r.id), ids)) edited.push(r);
      const done = entryOf(r)?.done;
      if (done !== undefined && done !== r.base.done) resolved.push([r, done]);
    }
    if (made.length + dropped.length + edited.length + resolved.length === 0) return { touched: [], bodyChanged: false };
    const main = this.session.main;
    let part = main.commentsPart;
    if (!part) {
      part = new CommentsPart();
      main.addTargetPart(part);
      info.tree = { TYPE_NAME: 'org_docx4j_wml.Comments', comment: [] };
    }
    if (!info.tree) return { touched: [], bodyChanged: false };
    const tree = info.tree;
    linkParents(tree, undefined);
    part.setContents(tree);
    const view = new Body(undefined, body, 'body', this.session.pkg);
    // The parts, as the engine's own views carry them: an anchored comment's, or a new comment's below.
    let parts: CommentParts | undefined = (await view.getComments())[0]?.parts;
    const views = new Map<string, Comment>();
    const viewsOver = (p: CommentParts): void => {
      for (const c of tree.comment ?? []) if (!views.has(String(c.id))) views.set(String(c.id), new Comment(c, p, view));
      for (const r of info.records.values()) {
        const parent = r.isNew ? null : parentOf(r);
        const child = views.get(r.id);
        const parentView = parent === null ? undefined : views.get(parent);
        if (child && parentView && !parentView.replies.includes(child)) parentView.addReply(child);
      }
    };
    if (parts) viewsOver(parts);
    const before = markersOf(body).length;
    // Parents before replies, each thread in the order of its ids.
    made.sort((a, b) => Number(parentOf(a) !== null) - Number(parentOf(b) !== null) || Number(a.id) - Number(b.id));
    for (const r of made) {
      this.exportStoryContent(r, store.comments.get(r.id), ids);
      const parent = parentOf(r);
      const parentView = parent === null ? undefined : views.get(parent);
      const entry = await addCommentEntry(view, {
        id: Number(r.id),
        content: (r.element.content ?? []) as Element<wml.P>[],
        author: { name: r.element.author },
        initials: r.element.initials ?? r.base.initials,
        date: r.base.date ?? new Date(),
        ...(parentView ? { parent: parentView } : {}),
      });
      // The engine's element is the comment now: the story keeps its projection, the element the entry.
      r.element = entry.element;
      views.set(r.id, entry);
      if (!parts) { parts = entry.parts; viewsOver(parts); }
      r.isNew = false;
      if (entryOf(r)?.done) resolved.push([r, true]);
    }
    for (const r of dropped) {
      const v = views.get(r.id);
      if (v && (tree.comment ?? []).includes(v.element)) await v.delete();
      // No view at all (no comment left anchored, none made): the element goes from the tree alone.
      else if (!v) tree.comment = (tree.comment ?? []).filter((c) => c !== r.element);
    }
    for (const [r, done] of resolved) { const v = views.get(r.id); if (v) v.resolved = done; }
    for (const r of [...made, ...edited]) keepParaId(r);
    for (const r of [...made, ...edited, ...resolved.map(([r]) => r)]) this.commentsWritten.add(r.id);
    return { touched: [part.partName.toString()], bodyChanged: markersOf(body).length !== before };
  }

  /**
   * The ids a save writes for the comments made in a room (ED-005 section 7.4), as notes'
   * (`noteIdsForFile`): each, in the order the document anchors them, the smallest id no other
   * comment in the part has, so Word 2010 never sees the room's random ones. Called after
   * `exportComments`, over the part as it will be written.
   */
  commentIdsForFile(doc: PMNode): CommentIds {
    const out: CommentIds = new Map();
    const tree = this.commentsLoaded?.tree;
    if (!tree) return out;
    const madeHere = (id: string): boolean => this.session.store.sources.has(commentSourceKey(id));
    const inPart = new Set((tree.comment ?? []).map((c) => String(c.id)));
    const used = new Set([...inPart].filter((id) => !madeHere(id)).map(Number));
    let next = 0;
    for (const id of commentAnchors(doc, this.session.ctx.ids)) {
      if (!madeHere(id) || !inPart.has(id) || out.has(Number(id))) continue;
      while (used.has(next)) next++;
      out.set(Number(id), next);
      used.add(next);
    }
    return out;
  }

  /** Write `ids` into the comments part's tree and the body's markers for a save; what it returns puts the room's ids back. */
  applyCommentIds(ids: CommentIds, body: wml.Body): () => void {
    const undo: (() => void)[] = [];
    if (ids.size === 0) return () => undefined;
    for (const c of this.commentsLoaded?.tree?.comment ?? []) {
      const to = ids.get(Number(c.id));
      if (to === undefined) continue;
      const from = c.id;
      c.id = to;
      undo.push(() => { c.id = from; });
    }
    for (const m of markersOf(body)) {
      const value = m.item.value as { id: number };
      const to = ids.get(Number(value.id));
      if (to === undefined) continue;
      const from = value.id;
      value.id = to;
      undo.push(() => { value.id = from; });
    }
    return () => { for (const f of [...undo].reverse()) f(); };
  }

  /**
   * The self-check over the comments the last save wrote or changed: each re-projected from the
   * saved bytes and compared with its live fragment, and its thread data (author, initials,
   * parent, done) with the reloaded part's.
   */
  private async selfCheckComments(reloaded: EditorDocument, ids: CommentIds): Promise<SelfCheckResult> {
    if (this.commentsWritten.size === 0) return { equal: true };
    await reloaded.stories.comments();
    const inFile = (id: string): string => String(ids.get(Number(id)) ?? id);
    for (const id of this.commentsWritten) {
      const mine = this.commentOf(id);
      if (!mine) continue;
      const other = reloaded.stories.commentOf(inFile(id));
      if (!other) return { equal: false, difference: { index: -1, live: 'present', reloaded: 'missing', story: mine.key } };
      const result = compareProjections(this.currentDoc(mine), other.doc);
      if (!result.equal) return { equal: false, difference: { ...result.difference!, story: mine.key } };
      const entry = readThreadEntry(this.session.store, id);
      const parent = entry?.parent !== undefined ? entry.parent : mine.base.parentId;
      const live = JSON.stringify([entry?.author ?? mine.base.author, entry?.initials ?? mine.base.initials, parent === null ? null : inFile(parent), entry?.done ?? mine.base.done]);
      const back = JSON.stringify([other.base.author, other.base.initials, other.base.parentId, other.base.done]);
      if (live !== back) return { equal: false, difference: { index: -1, live, reloaded: back, story: mine.key } };
    }
    return { equal: true };
  }

  // ---------------------------------------------------------------- headers and footers

  /** Every header and footer part of the main document, by relationship id. */
  headers(): Promise<HeaderFooter[]> {
    return (this.headersCache ??= this.readHeaders().then((h) => { this.headersLoaded = h; this.notify(); return h; }));
  }

  /**
   * A header or footer part made in the session (a page number's footer, ED-005 section 12.95; a
   * peer's arriving through the part sync): projected and listed with the ones read, whether the
   * read has happened, is under way, or is still to come (then the read finds the part itself).
   */
  headerAdded(kind: 'header' | 'footer', part: HeaderPart | FooterPart, rId: string, element: wml.Hdr | wml.Ftr): HeaderFooter {
    const record: HeaderFooter = { rId, kind, part, element, ...projectStory(this.session, element.content as Element[] | undefined, `${kind}:${part.partName.toString()}`) };
    const add = (list: HeaderFooter[]): HeaderFooter[] => { if (!list.some((h) => h.kind === kind && h.rId === rId)) list.push(record); return list; };
    if (this.headersLoaded) { add(this.headersLoaded); this.notify(); }
    else if (this.headersCache) this.headersCache = this.headersCache.then(add);
    return record;
  }

  /** Whether the document has different odd and even pages (`w:evenAndOddHeaders`). */
  async evenAndOddHeaders(): Promise<boolean> {
    const settings = await this.settings();
    return settings?.evenAndOddHeaders !== undefined && settings.evenAndOddHeaders.val !== false;
  }

  /** Blocks projected from elements and put at a story's end through its fragment, one local transaction (a page number's paragraph). */
  appendBlocks(story: Story, elements: Element[]): void {
    const blocks = projectBlocks(elements, story.ctx, story.key);
    const doc = this.currentDoc(story);
    this.fragmentOf(story, true);
    this.replace(story, doc.type.create(doc.attrs, doc.content.append(Fragment.from(blocks))));
  }

  /** A header or footer by part name, once read. */
  headerByPartName(partName: string): HeaderFooter | undefined {
    return this.headersLoaded?.find((h) => h.part.partName.toString() === partName);
  }

  private async readHeaders(): Promise<HeaderFooter[]> {
    const main = this.session.main;
    const rp = main.relationshipsPart;
    if (!rp) return [];
    const out: HeaderFooter[] = [];
    for (const kind of ['header', 'footer'] as const) {
      for (const rel of rp.getRelationshipsByType(kind === 'header' ? Namespaces.HEADER : Namespaces.FOOTER)) {
        const part = rp.getPart(rel) as HeaderPart | FooterPart | undefined;
        if (!part) continue;
        let element: wml.Hdr | wml.Ftr;
        try { element = await part.readContents(); } catch { continue; }
        out.push({ rId: rel.id, kind, part, element, ...projectStory(this.session, element.content as Element[] | undefined, `${kind}:${part.partName.toString()}`) });
      }
    }
    return out;
  }

  private settings(): Promise<wml.CTSettings | undefined> {
    return (this.settingsCache ??= (async () => {
      const part = this.session.main.documentSettingsPart;
      if (!part) return undefined;
      try { return await part.readContents(); } catch { return undefined; }
    })());
  }

  /**
   * The sections in document order: the paragraph-level `sectPr`s, then the final one. Over the
   * document given, or the session's, which is read from the store afresh each time: a caller
   * with a document in hand (a view's state) passes it, and one asking per page passes the
   * result on to `headerFooterFor`.
   */
  sections(doc: PMNode = this.session.doc): (wml.SectPr | null)[] {
    const out: (wml.SectPr | null)[] = [];
    doc.descendants((node) => {
      if (node.type.name !== 'paragraph') return true;
      const s = node.attrs['sectPr'] as wml.SectPr | null;
      if (s) out.push(s);
      return false;
    });
    out.push((doc.attrs['finalSectPr'] as wml.SectPr | null) ?? null);
    return out;
  }

  /**
   * The header and footer a page shows (ED-001 Appendix D.2), as Word decides: the section's
   * `first` reference on its first page when `w:titlePg` is set, `even` on even pages when the
   * settings part sets `w:evenAndOddHeaders`, else `default`; a section without a reference of
   * the type inherits the previous section's (Word's "same as previous"); the first has none.
   * `page` is the number the page shows (from `w:pgNumType`, ED-003 section 6.4), whose parity
   * picks the even variant.
   */
  async headerFooterFor(sectionIndex: number, page: number, firstPageOfSection: boolean, sections: (wml.SectPr | null)[] = this.sections()): Promise<PageStories> {
    const [headers, settings] = await Promise.all([this.headers(), this.settings()]);
    const evenAndOdd = settings?.evenAndOddHeaders !== undefined && settings.evenAndOddHeaders.val !== false;
    const section = sections[sectionIndex] ?? null;
    const titlePg = section?.titlePg !== undefined && section.titlePg.val !== false;
    const variant: HeaderVariant = firstPageOfSection && titlePg ? 'first' : evenAndOdd && page % 2 === 0 ? 'even' : 'default';
    const find = (kind: 'header' | 'footer'): HeaderFooter | undefined => {
      const typeName = kind === 'header' ? 'org_docx4j_wml.HeaderReference' : 'org_docx4j_wml.FooterReference';
      for (let i = sectionIndex; i >= 0; i--) {
        const refs = (sections[i]?.egHdrFtrReferences ?? []) as (wml.HeaderReference | wml.FooterReference)[];
        const ref = refs.find((r) => r.TYPE_NAME === typeName && r.type === variant);
        if (ref) return headers.find((h) => h.kind === kind && h.rId === ref.id);
        // A section that names any reference of this kind but not this variant falls back to its default.
        const fallback = variant !== 'default' ? refs.find((r) => r.TYPE_NAME === typeName && r.type === 'default') : undefined;
        if (fallback) return headers.find((h) => h.kind === kind && h.rId === fallback.id);
      }
      return undefined;
    };
    const out: PageStories = { variant };
    const header = find('header');
    const footer = find('footer');
    if (header) out.header = header;
    if (footer) out.footer = footer;
    return out;
  }
}

/**
 * Each comment's `w16cex:dateUtc`, by the paragraph id `w16cid:commentId` pairs its durable id with,
 * as the engine's `Comment.creationDate` reads them (its `readDatesUtc`, which it does not export):
 * from the parts' private trees, which leaves them untouched.
 */
async function commentDatesUtc(main: MainDocumentPart): Promise<Map<string, Date>> {
  const out = new Map<string, Date>();
  if (!main.commentsIdsPart || !main.commentsExtensiblePart) return out;
  let ids: { commentId?: { durableId?: string; paraId?: string }[] };
  let extensible: { commentExtensible?: { durableId?: string; dateUtc?: wml.XmlCalendar }[] };
  try {
    ids = await main.commentsIdsPart.readContents() as typeof ids;
    extensible = await main.commentsExtensiblePart.readContents() as typeof extensible;
  } catch {
    return out;
  }
  const paraIdOf = new Map<string, string>();
  for (const e of ids.commentId ?? []) if (e.durableId && e.paraId) paraIdOf.set(e.durableId.toUpperCase(), e.paraId.toUpperCase());
  for (const e of extensible.commentExtensible ?? []) {
    const paraId = e.durableId ? paraIdOf.get(e.durableId.toUpperCase()) : undefined;
    const date = revisionDateOf({ dateUtc: e.dateUtc });
    if (paraId && date && !Number.isNaN(date.getTime())) out.set(paraId, date);
  }
  return out;
}

/**
 * When a comment was made, as Word means it (core-ts 0.3.0, CR-002 section 29; the engine's
 * `Comment.creationDate`): its `w16cex:dateUtc` when the document has one for it, else `w:date` read
 * as the local wall clock, its `Z` notwithstanding (ED-005 section 12.37: read as UTC, a comment an
 * agent's tree run made was dated by the time zone's offset on the other peers).
 */
function commentDate(c: wml.Comments.Comment, datesUtc: ReadonlyMap<string, Date>): Date | undefined {
  for (const el of [...(c.content ?? [])].reverse()) {
    const paraId = typeNameOf(el) === 'org_docx4j_wml.P' ? (el.value as wml.P).paraId : undefined;
    const utc = paraId ? datesUtc.get(paraId.toUpperCase()) : undefined;
    if (utc) return utc;
  }
  return revisionDateOf({ date: c.date });
}

// SPDX-License-Identifier: GPL-3.0-or-later
//
// The editor's own clipboard (ED-010). Copy writes the slice's JSON on the clipboard's first
// element beside ProseMirror's `data-pm-slice`, stamped with the session's token; a paste whose
// stamp is this session's rebuilds the slice from it under the model's rules (`pastedSlice`),
// the copies it needs made first (a comment copied whole made again, a note copied, a picture
// registered afresh), and dispatches it as one transaction through the surface, so the tracker
// and the control rules run over it as over any local edit. A stamp from another session falls
// through to the HTML, as before. A dragged copy goes through the same rules in `transformPasted`
// (its copies the synchronous ones); a dragged move is left as it is, since nothing is duplicated.
import { DOMSerializer, Slice, type Node as PMNode, type Schema } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { deepCopy, linkParents, runOf, type Element } from '@docx4j/core-ts';
import {
  pastedSlice, closedAtMark, commentLengthsIn, commentLengthIn, allocateSdtId, inlineOf, newCommentReference, newNoteReference, storyOf, NOTE_STYLES, PASTE_META,
  CLIPBOARD_SLICE_ATTR, CLIPBOARD_SESSION_ATTR, type EditorDocument, type PasteRules, type NoteKind, type CommentRecord,
} from '@docx4j/editor-model';

// ---------------------------------------------------------------- copy

/**
 * ProseMirror's serializer with the editor's two attributes on the fragment's first element. The
 * slice itself comes from `transformCopied`, which ProseMirror calls just before.
 */
export class EditorClipboardSerializer extends DOMSerializer {
  constructor(schema: Schema, private readonly token: string, private readonly copied: () => Slice | null) {
    const base = DOMSerializer.fromSchema(schema);
    super(base.nodes, base.marks);
  }

  override serializeFragment(fragment: Parameters<DOMSerializer['serializeFragment']>[0], options?: Parameters<DOMSerializer['serializeFragment']>[1], target?: Parameters<DOMSerializer['serializeFragment']>[2]): ReturnType<DOMSerializer['serializeFragment']> {
    const dom = super.serializeFragment(fragment, options, target);
    const slice = this.copied();
    const first = Array.from(dom.childNodes).find((n): n is HTMLElement => n.nodeType === 1);
    if (slice && first) {
      first.setAttribute(CLIPBOARD_SLICE_ATTR, JSON.stringify(slice.toJSON()));
      first.setAttribute(CLIPBOARD_SESSION_ATTR, this.token);
    }
    return dom;
  }
}

/** The slice JSON an editor clipboard carries, when its stamp is `token`'s; null otherwise (another session's, or not the editor's). */
export function editorSliceOf(html: string, token: string): unknown | null {
  if (!html.includes(CLIPBOARD_SLICE_ATTR)) return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const el = doc.querySelector(`[${CLIPBOARD_SLICE_ATTR}]`);
  if (!el || el.getAttribute(CLIPBOARD_SESSION_ATTR) !== token) return null;
  try { return JSON.parse(el.getAttribute(CLIPBOARD_SLICE_ATTR) ?? ''); } catch { return null; }
}

// ---------------------------------------------------------------- paste

/** The rules' context for a view: the mode, the story, a control id, and the registry for the atoms' elements. */
function baseRules(view: EditorView, session: EditorDocument): PasteRules {
  return {
    tracking: session.tracking() !== undefined,
    story: storyOf(view.state) !== undefined,
    controlId: () => allocateSdtId(new Set()),
    bookmarkOf: bookmarkRule(view, session),
  };
}

/**
 * A bookmark atom's id and name from its element, and whether a bookmark of that name is in the
 * document now (one cut is not, so it comes across; one copied is, so it is dropped, as Word
 * does both). The names are the inline markers' in this view's document; a marker between blocks
 * is side data the walk does not see.
 */
function bookmarkRule(view: EditorView, session: EditorDocument): PasteRules['bookmarkOf'] {
  const valueOf = (node: PMNode): { id?: number | string; name?: string } | undefined => session.ctx.ids.get(String(node.attrs['sourceId']))?.value as { id?: number | string; name?: string } | undefined;
  const names = new Set<string>();
  view.state.doc.descendants((n) => {
    if (n.type.name === 'opaqueInline' && n.attrs['kind'] === 'bookmarkStart') { const name = valueOf(n)?.name; if (name !== undefined) names.add(name); }
    return true;
  });
  return (node) => {
    const v = valueOf(node);
    if (!v || v.id === undefined) return undefined;
    return { id: String(v.id), name: v.name ?? '', taken: v.name !== undefined && names.has(v.name) };
  };
}

/** The comment a reference atom stands for, by its element's id. */
function referencedCommentId(session: EditorDocument, node: PMNode): string | undefined {
  if (node.type.name !== 'opaqueInline' || node.attrs['kind'] !== 'commentReference') return undefined;
  const el = session.ctx.ids.get(String(node.attrs['sourceId']));
  const id = (el?.value as { id?: number | string } | undefined)?.id;
  return id === undefined ? undefined : String(id);
}

/**
 * A picture's drawing copied and registered afresh, with a fresh `wp:docPr` id and name, as Word
 * gives a pasted picture (ED-010 section 9.1, case 6; the image part and relationship shared).
 * The copy's run goes to `sources`, so another session resolves it. Synchronous, so a dragged
 * copy gets it too.
 */
function copiedPicture(session: EditorDocument, node: PMNode): PMNode | null {
  const schema = node.type.schema;
  const source = session.ctx.ids.get(String(node.attrs['sourceId']));
  if (!source) return null;
  const copy = deepCopy(source) as Element;
  const inline = inlineOf(copy) as { docPr?: { id?: number; name?: string } } | undefined;
  if (inline?.docPr) {
    const id = 2 ** 20 + Math.floor(Math.random() * (2 ** 31 - 2 ** 20));
    inline.docPr.id = id;
    inline.docPr.name = `Picture ${id}`;
  }
  const sourceRun = session.ctx.ids.get(String(node.attrs['runSource']));
  const rPr = (sourceRun?.value as { rPr?: unknown } | undefined)?.rPr;
  const run = runOf([copy], rPr ? (deepCopy(rPr as object) as never) : undefined);
  linkParents(run.value, undefined);
  const sourceId = session.ctx.ids.idOf(copy);
  const runSource = session.ctx.ids.idOf(run);
  void session.writeInlineSource(run);
  return schema.nodes['image']!.create({ ...node.attrs, sourceId, runSource });
}

/**
 * A note copied: a new note of the kind holding a copy of the source's content, its own
 * reference mark in place of the source's (ED-010 section 9.1, case 7), and a reference run
 * registered for the body. Null when the source is not known here.
 */
async function copiedNote(session: EditorDocument, node: PMNode): Promise<PMNode | null> {
  const kind = node.attrs['kind'] as NoteKind;
  const source = session.stories.noteOf(kind, Number(node.attrs['footnoteId']));
  if (!source) return null;
  await session.ensureNotesPart(kind);
  await session.ensureStyles([NOTE_STYLES[kind].text, NOTE_STYLES[kind].reference]);
  const note = await session.stories.newNote(kind, '');
  const sourceDoc = session.stories.currentDoc(source);
  const noteDoc = session.stories.currentDoc(note);
  const marks = new Set(['footnoteRef', 'endnoteRef']);
  let ownMark: PMNode | null = null;
  noteDoc.descendants((n) => { if (!ownMark && n.type.name === 'opaqueInline' && marks.has(String(n.attrs['kind']))) ownMark = n; return !ownMark; });
  let placed = false;
  const copied = pastedSlice(new Slice(sourceDoc.content, 0, 0), {
    tracking: false, story: true,
    atom: (n) => {
      if (n.type.name === 'opaqueInline' && marks.has(String(n.attrs['kind']))) {
        if (placed) return null;
        placed = true;
        return ownMark;
      }
      return undefined;
    },
  });
  session.stories.replace(note, noteDoc.type.create(noteDoc.attrs, copied.content));
  const { run, ...reference } = newNoteReference(session, kind, note.id, undefined);
  await session.writeInlineSource(run);
  return node.type.schema.nodes['footnoteRef']!.create({ ...node.attrs, ...reference, footnoteId: note.id });
}

/**
 * The comments a slice carries whole, made again (ED-010 section 9.1, case 1: the original's author,
 * date, text and replies, a new id); the map from the old ids to the new, and a reference atom's
 * replacement per old id. A comment carried in part is not in the map, so its mark goes (case 2).
 */
async function copiedComments(view: EditorView, session: EditorDocument, slice: Slice): Promise<{ ids: Map<string, string>; references: Map<string, { sourceId: string; runSource: string }> }> {
  const ids = new Map<string, string>();
  const references = new Map<string, { sourceId: string; runSource: string }>();
  const lengths = commentLengthsIn(slice.content);
  if (lengths.size === 0) return { ids, references };
  const threads = await session.stories.comments(view.state.doc);
  const records = new Map<string, CommentRecord>();
  const collect = (list: typeof threads): void => { for (const t of list) { records.set(t.id, t.record); collect(t.replies); } };
  collect(threads);
  const whole = [...lengths].filter(([id, length]) => length > 0 && length === commentLengthIn(view.state.doc, id) && records.has(id)).map(([id]) => id);
  if (whole.length === 0) return { ids, references };
  await session.ensureCommentStyles();
  const copy = async (id: string, parent: string | null): Promise<void> => {
    const record = records.get(id)!;
    const made = await session.stories.newComment({ author: record.base.author, initials: record.base.initials, text: '', ...(record.base.date ? { date: record.base.date } : {}), ...(parent ? { parent } : {}) });
    session.stories.replace(made, session.stories.currentDoc(record));
    const { run, ...reference } = newCommentReference(session, Number(made.id), undefined);
    await session.writeInlineSource(run);
    ids.set(id, made.id);
    references.set(id, reference);
  };
  // parents first, so a reply copied with its parent joins the copy's thread
  for (const id of whole) if (records.get(id)!.base.parentId === null) await copy(id, null);
  for (const id of whole) {
    const parentId = records.get(id)!.base.parentId;
    if (parentId !== null && ids.has(parentId)) await copy(id, ids.get(parentId)!);
  }
  return { ids, references };
}

/**
 * The editor's own clipboard pasted: the slice rebuilt from its JSON, the copies made, the rules
 * applied, one transaction through the surface. False when the JSON is not a slice of this schema
 * (the caller then pastes the HTML).
 */
export async function pasteSlice(view: EditorView, session: EditorDocument, json: unknown): Promise<boolean> {
  let slice: Slice;
  try { slice = closedAtMark(Slice.fromJSON(view.state.schema, json)); } catch { return false; }
  const rules = baseRules(view, session);
  const comments = rules.story ? { ids: new Map<string, string>(), references: new Map<string, { sourceId: string; runSource: string }>() } : await copiedComments(view, session, slice);
  const notes = new Map<PMNode, PMNode | null>();
  if (!rules.story) {
    const refs: PMNode[] = [];
    slice.content.descendants((n) => { if (n.type.name === 'footnoteRef') refs.push(n); return true; });
    for (const ref of refs) notes.set(ref, await copiedNote(session, ref));
  }
  if ((view as { isDestroyed?: boolean }).isDestroyed) return true;
  const out = pastedSlice(slice, {
    ...rules,
    commentId: (id) => comments.ids.get(id),
    atom: (node) => {
      switch (node.type.name) {
        case 'image': return copiedPicture(session, node);
        case 'footnoteRef': return rules.story ? null : notes.get(node) ?? null;
        case 'opaqueInline': {
          const id = referencedCommentId(session, node);
          if (id === undefined) return undefined;
          const reference = comments.references.get(id);
          return reference ? node.type.create({ ...node.attrs, ...reference }) : null;
        }
        default: return undefined;
      }
    },
  });
  const tr = view.state.tr.replaceSelection(out).setMeta(PASTE_META, true);
  view.dispatch(tr.scrollIntoView());
  return true;
}

/**
 * ProseMirror's `transformPasted`: a dragged copy, and any slice ProseMirror parsed itself, go
 * through the rules with the synchronous copies (a picture); a comment's reference and a note
 * reference are dropped there, since their copies are made before a transaction, not inside
 * one. A dragged move is left whole: nothing is duplicated.
 */
export function transformPastedSlice(slice: Slice, view: EditorView, session: EditorDocument): Slice {
  if (view.dragging?.move) return slice;
  const rules = baseRules(view, session);
  // A dragged copy's slice is a selection's, closed at a paragraph mark; a parsed paste's is not.
  return pastedSlice(view.dragging ? closedAtMark(slice) : slice, {
    ...rules,
    commentId: () => undefined,
    atom: (node) => {
      switch (node.type.name) {
        case 'image': return copiedPicture(session, node);
        case 'footnoteRef': return null;
        case 'opaqueInline': return referencedCommentId(session, node) === undefined ? undefined : null;
        default: return undefined;
      }
    },
  });
}

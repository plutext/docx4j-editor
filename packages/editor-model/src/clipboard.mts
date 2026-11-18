// SPDX-License-Identifier: GPL-3.0-or-later
//
// Copy and paste within the editor (ED-010): the rules a pasted slice goes through, pure over the
// slice, so the view's paste, a dragged copy and the tests share them. The view carries the
// slice's JSON on the clipboard beside ProseMirror's HTML, stamped with the session's token, and
// rebuilds it here when the stamp is its own (ED-010 section 3).
//
// The rules (section 3.3, and section 9.1 for what Word 365 measured):
// - a paragraph drops its `sourceId` and names it as `copiedFrom`, Enter's rule (ED-005 section
//   12.102), so the export takes the source's whole `w:pPr` as the copy's base; its section is
//   dropped;
// - a table, a row and a cell drop their `sourceId` (their properties are attributes) and a row its
//   row-level control stack, whose group is not copied;
// - a control gets a fresh id, as Word gives a pasted control one;
// - a link drops the holder it came from and keeps its address, so the export makes a holder;
// - with tracking on the slice is taken as accepted first (deleted text out, insertion and
//   formatting marks off, a deleted row out), which the tracker then marks this author's
//   insertion, as Word pastes (check: ED-010 section 9.1, case 4); with tracking off the source
//   author's revisions are kept, as Word keeps them (case 3);
// - a bookmark whose name is still in the document is dropped, as Word's own paste drops it (case
//   5, a copy: the original keeps the name); one whose name is free, both its markers in the slice,
//   comes across, as Word moves a bookmark with what is cut and pasted (its help on bookmarks); with
//   no `bookmarkOf` every bookmark is dropped;
// - a slice of the editor's own that ends with a paragraph mark (its last child an empty open
//   textblock, the selection having run to the next paragraph's start: the selection bar's,
//   Shift+Down's) is closed by its caller through `closedAtMark`, so the pasted paragraphs stand
//   whole with their own properties where ProseMirror would join the first into the target and give
//   it the target's (Jason, 2026-10-29: a cut heading lost its style). An HTML or plain-text paste
//   is not closed: its trailing empty paragraph is the target's own mark, as Word leaves it;
// - a comment's mark takes the id the caller gives it (a comment copied whole is made again, case
//   1) or goes (copied in part, case 2); the reference atom likewise;
// - an image, a note reference and any other atom are the caller's to replace (a copy registered
//   afresh, a new note) or drop, through `atom`;
// - into a story (a note, a header) tables, comments and revisions have no place and go.
import { Fragment, Slice, type Node as PMNode, type Mark, type Schema } from 'prosemirror-model';
import { withSdtId } from './sdtpr.mjs';
import type { Revision } from './schema.mjs';

/** What a paste's rules need to know, and what the caller supplies for the atoms it can copy. */
export interface PasteRules {
  /** Tracking on: the slice is taken as accepted first. */
  tracking: boolean;
  /** Into a story (a note, a header): tables, comments and revision marks go. */
  story: boolean;
  /** A comment's new id for an old one; undefined drops the mark (and the reference atom). */
  commentId?: (id: string) => string | undefined;
  /** An atom's replacement: `null` drops it, `undefined` keeps it as it is. */
  atom?: (node: PMNode) => PMNode | null | undefined;
  /** A fresh control id for a pasted control. */
  controlId?: () => number;
  /** A bookmark atom's id and name, and whether the name is already in the document; undefined for any other atom. */
  bookmarkOf?: (node: PMNode) => { id: string; name: string; taken: boolean } | undefined;
}

const BOOKMARKS = new Set(['bookmarkStart', 'bookmarkEnd']);

/** The slice as the rules leave it; the same openness, unless a top-level block went, when it is opened as far as it goes. */
export function pastedSlice(slice: Slice, rules: PasteRules): Slice {
  const content = fragmentOf(slice.content, { ...rules, kept: keptBookmarks(slice, rules) });
  if (content.childCount !== slice.content.childCount) return Slice.maxOpen(content);
  return new Slice(content, slice.openStart, slice.openEnd);
}

/**
 * A slice ending with a paragraph mark, closed: the selection ran from inside a paragraph to the
 * next paragraph's start, so the last child is an empty open textblock standing for that mark. It
 * goes, and the slice is closed at both ends, so the paragraphs it holds paste whole with their
 * own properties, as Word pastes paragraphs whose marks were selected. Any other slice is as it was.
 * For the editor's own clipboard and a dragged copy, whose slices come from a selection; not for an
 * HTML or plain-text paste, where a trailing empty paragraph is the target's mark.
 */
export function closedAtMark(slice: Slice): Slice {
  if (slice.openStart > 1 || slice.openEnd !== 1 || slice.content.childCount === 0) return slice;
  const last = slice.content.lastChild!;
  if (!last.isTextblock || last.content.size !== 0) return slice;
  const content = slice.content.cut(0, slice.content.size - last.nodeSize);
  if (content.childCount === 0) return slice;
  return new Slice(content, 0, 0);
}

/** The ids of the bookmarks the slice keeps: both markers in it, the name not in the document. */
function keptBookmarks(slice: Slice, rules: PasteRules): Set<string> {
  const out = new Set<string>();
  if (!rules.bookmarkOf) return out;
  const starts = new Map<string, string>();
  const ends = new Set<string>();
  slice.content.descendants((node) => {
    if (node.type.name !== 'opaqueInline' || !BOOKMARKS.has(String(node.attrs['kind']))) return true;
    const b = rules.bookmarkOf!(node);
    if (!b) return false;
    if (node.attrs['kind'] === 'bookmarkStart') { if (!b.taken) starts.set(b.id, b.name); } else ends.add(b.id);
    return false;
  });
  for (const id of starts.keys()) if (ends.has(id)) out.add(id);
  return out;
}

type Rules = PasteRules & { kept: Set<string> };

function fragmentOf(fragment: Fragment, rules: Rules): Fragment {
  const out: PMNode[] = [];
  fragment.forEach((node) => {
    const mapped = nodeOf(node, rules);
    if (mapped) out.push(mapped);
  });
  return Fragment.from(out);
}

/** The revision an atom, a mark or a row keeps: with tracking on none (a deleted one is dropped by the caller); off, as it was. */
function revisionKept(revision: Revision | null, rules: Rules): { drop: boolean; value: Revision | null } {
  if (!revision) return { drop: false, value: null };
  if (!rules.tracking && !rules.story) return { drop: false, value: revision };
  return { drop: revision.deletion !== undefined, value: null };
}

function nodeOf(node: PMNode, rules: Rules): PMNode | null {
  const schema: Schema = node.type.schema;
  switch (node.type.name) {
    case 'text': {
      const marks = marksOf(node.marks, rules);
      return marks === null ? null : node.mark(marks);
    }
    case 'paragraph': {
      const attrs = node.attrs;
      const kept = revisionKept(attrs['markRevision'] as Revision | null, rules);
      return schema.nodes['paragraph']!.create({
        ...attrs,
        sourceId: null,
        copiedFrom: (attrs['sourceId'] as string | null) ?? (attrs['copiedFrom'] as string | null),
        sectPr: null,
        // a deleted mark's paragraph joins the next with tracking on: the mark is dropped by the
        // caller's view of the slice only through the tracker; here the revision alone goes
        markRevision: kept.value,
        pPrChange: rules.tracking || rules.story ? null : attrs['pPrChange'],
      }, fragmentOf(node.content, rules));
    }
    case 'table': {
      if (rules.story) return null;
      return schema.nodes['table']!.create({ ...node.attrs, sourceId: null }, fragmentOf(node.content, rules));
    }
    case 'table_row': {
      const kept = revisionKept(node.attrs['revision'] as Revision | null, rules);
      if (kept.drop) return null;
      return schema.nodes['table_row']!.create({ ...node.attrs, sourceId: null, sdt: null, revision: kept.value }, fragmentOf(node.content, rules));
    }
    case 'table_cell':
      return schema.nodes['table_cell']!.create({ ...node.attrs, sourceId: null }, fragmentOf(node.content, rules));
    case 'sdtBlock':
    case 'sdtInline': {
      const sdtPr = node.attrs['sdtPr'];
      return node.type.create({
        ...node.attrs,
        sourceId: null,
        sdtPr: sdtPr && rules.controlId ? withSdtId(sdtPr, rules.controlId()) : sdtPr,
      }, fragmentOf(node.content, rules));
    }
    default: {
      if (!node.isAtom && !node.isLeaf) return node.copy(fragmentOf(node.content, rules));
      if (node.type.name === 'opaqueInline' && BOOKMARKS.has(String(node.attrs['kind']))) {
        const b = rules.bookmarkOf?.(node);
        if (!b || !rules.kept.has(b.id)) return null;
      }
      const kept = 'revision' in node.attrs ? revisionKept(node.attrs['revision'] as Revision | null, rules) : { drop: false, value: null };
      if (kept.drop) return null;
      let out: PMNode | null | undefined = node;
      if (rules.atom && node.isInline) {
        out = rules.atom(node);
        if (out === null) return null;
        if (out === undefined) out = node;
      }
      const marks = marksOf(out.marks, rules);
      if (marks === null) return null;
      const withRevision = 'revision' in out.attrs && out.attrs['revision'] !== kept.value ? out.type.create({ ...out.attrs, revision: kept.value }, out.content, marks) : out.mark(marks);
      return withRevision;
    }
  }
}

/** The marks a pasted text or atom keeps; null when the content goes (deleted text, tracking on). */
function marksOf(marks: readonly Mark[], rules: Rules): Mark[] | null {
  const out: Mark[] = [];
  for (const mark of marks) {
    switch (mark.type.name) {
      case 'deletion':
        if (rules.tracking || rules.story) return null;
        out.push(mark);
        break;
      case 'insertion':
      case 'formatChange':
        if (!(rules.tracking || rules.story)) out.push(mark);
        break;
      case 'comment': {
        if (rules.story) break;
        const id = rules.commentId?.(String(mark.attrs['id']));
        if (id !== undefined) out.push(id === mark.attrs['id'] ? mark : mark.type.create({ ...mark.attrs, id }));
        break;
      }
      case 'link':
        out.push(mark.attrs['sourceId'] === null ? mark : mark.type.create({ ...mark.attrs, sourceId: null }));
        break;
      default:
        out.push(mark);
    }
  }
  return out;
}

/** The text each comment marks in a fragment, by id (UTF-16 units), for the whole-or-part test (ED-010 section 9.1, cases 1 and 2). */
export function commentLengthsIn(fragment: Fragment): Map<string, number> {
  const out = new Map<string, number>();
  fragment.descendants((node) => {
    if (!node.isText) return true;
    for (const mark of node.marks) if (mark.type.name === 'comment') out.set(String(mark.attrs['id']), (out.get(String(mark.attrs['id'])) ?? 0) + node.nodeSize);
    return true;
  });
  return out;
}

/** The text a comment marks in a document (UTF-16 units). */
export function commentLengthIn(doc: PMNode, id: string): number {
  let length = 0;
  doc.descendants((node) => {
    if (node.isText && node.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === id)) length += node.nodeSize;
    return true;
  });
  return length;
}

/** The editor's attributes on the clipboard's first element: the slice as JSON, and the session that wrote it. */
export const CLIPBOARD_SLICE_ATTR = 'data-docx4j-slice';
export const CLIPBOARD_SESSION_ATTR = 'data-docx4j-session';

/** A session's clipboard token: random, so a paste from another tab, another document or a peer falls through to the HTML. */
export function clipboardToken(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

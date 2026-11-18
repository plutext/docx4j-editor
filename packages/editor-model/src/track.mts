// SPDX-License-Identifier: GPL-3.0-or-later
//
// The tracker (ED-005 section 5.4, step C3): a local transaction, with tracking on, rewritten
// before it is applied, by one function the view's dispatch and the headless runner both call,
// so a command behaves the same in a test and on screen. The rewrite appends steps to the
// transaction rather than building another: deleted content is put back with the deletion mark,
// inserted content takes the insertion mark, a formatting change takes `formatChange` and a
// paragraph's properties `pPrChange`. y-prosemirror writes the difference between the documents
// to Yjs, so what was deleted and put back reaches the room as a mark on the same text, and a
// peer's concurrent edit inside it survives.
//
// The replaced ranges come from prosemirror-changeset, whose tokens here make every character
// distinct (a word typed over is a deletion and an insertion whole, as Word shows it), a block
// equal to one of its type (a paragraph's attributes changed are not content replaced) and an atom
// equal to one that differs from it in its formatting alone (a tab made bold is not replaced).
// A paragraph mark a split names as new (`NEW_MARK_META`) is equal to none, so the mark the
// tracker takes as inserted is Word's, at the caret (core-ts check 29).
import { ChangeSet, type TokenEncoder } from 'prosemirror-changeset';
import { Selection, type EditorState, type Transaction } from 'prosemirror-state';
import { AddMarkStep, AttrStep, RemoveMarkStep, ReplaceAroundStep, canJoin, type StepMap } from 'prosemirror-transform';
import type { Mark, MarkType, Node as PMNode } from 'prosemirror-model';
import { ySyncPluginKey } from 'y-prosemirror';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Element } from '@docx4j/core-ts';
import { FORMATTING_MARKS, type ParagraphAttrs, type Revision, type RevisionInfo } from './schema.mjs';
import { runPropertiesFor, paragraphPropertiesOf } from './export.mjs';
import { formatChangeBeforeRPr } from './revisions.mjs';
import { plainEqual, revive, sameValue, toPlain } from './json.mjs';
import { commandNamed, sessionOf, storyOf, COMMAND_META, NEW_MARK_META } from './commands/spec.mjs';
import { paragraphAttrsOf, projectedParagraphAttrs, type ProjectionContext } from './import.mjs';
import { writeExcludedCharacters } from './excluded.mjs';
import { keepBookmarkMarkers } from './bookmarks.mjs';
import { keepCommentsWhole } from './commentrange.mjs';

/** A transaction the tracker leaves as it is: accept and reject (step C4), and anything else that must not be recorded. */
export const UNTRACKED_META = 'docx4j-untracked';
/**
 * A paste of the editor's own clipboard (ED-010): with tracking off, what it inserts keeps the
 * source author's revisions, as Word keeps them (section 9.1, case 3), so the transaction is not
 * shorn of them as typed text is; with tracking on the slice was taken as accepted before the
 * transaction, and the tracker marks it as any insertion.
 */
export const PASTE_META = 'docx4j-paste-slice';

/** Who a tracked change made now is by, and when: to the minute, Word's resolution (section 5.4). */
export interface Tracking {
  author: string;
  date: string;
}

/** Now to the minute, as `w:date` holds it: a typing run is one change, as in Word (section 5.4). */
export function trackingDate(now: Date = new Date()): string {
  const d = new Date(now.getTime());
  d.setUTCSeconds(0, 0);
  return d.toISOString().replace(/\.000Z$/, 'Z');
}

/** What became of a local transaction: recorded as tracked, or left untracked for a reason the status line gives. */
export interface Recorded {
  tr: Transaction;
  /** Set when tracking is on and the transaction ran untracked: the command's name, or `story` for a note or header. */
  untracked?: string;
}

/**
 * A transaction as this session records it (section 5.4). A remote one (y-prosemirror's, undo
 * and redo included) and one marked `UNTRACKED_META` pass as they are. With tracking on, a
 * story's (a note, a header: E1's projection has no marks to hold a change) and an untracked
 * command's run untracked and say so; the rest are tracked. With tracking off, what the
 * transaction made sheds the revision it inherited from where it was made (ED-005 section 12.5
 * item 11, section 12.16).
 */
export function recordTransaction(tr: Transaction, state: EditorState): Recorded {
  if (!tr.docChanged || (tr.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean } | undefined)?.isChangeOrigin) return { tr };
  writeExcludedCharacters(tr);
  // A comment stays one range, so what this inserts inside one takes it (ED-005 section 12.116).
  keepCommentsWhole(tr);
  const recorded = recordLocal(tr, state);
  // After the tracker, which keeps what it deletes: a bookmark's marker removed with its partner left goes back (ED-005 section 12.115).
  const ids = sessionOf(state)?.ctx.ids;
  if (ids) keepBookmarkMarkers(recorded.tr, (sourceId) => bookmarkIdOf(ids.get(sourceId)));
  return recorded;
}

/** A bookmark marker's `w:id`, from its element. */
function bookmarkIdOf(el: Element | undefined): string | undefined {
  const id = (el?.value as { id?: unknown } | undefined)?.id;
  return id === undefined || id === null ? undefined : String(id);
}

function recordLocal(tr: Transaction, state: EditorState): Recorded {
  if (tr.getMeta(UNTRACKED_META)) return { tr };
  const session = sessionOf(state);
  if (!session) return { tr };
  const tracking = session.tracking();
  const name = tr.getMeta(COMMAND_META) as string | undefined;
  if (!tracking) return { tr: tr.getMeta(PASTE_META) ? tr : withoutInheritedRevisions(tr) };
  if (storyOf(state)) return { tr, untracked: 'story' };
  if (name !== undefined && commandNamed(name)?.tracking === 'untracked') return { tr, untracked: name };
  return { tr: trackTransaction(tr, state, tracking, session.ctx) };
}

// ---------------------------------------------------------------------------------------------
// Tracked
// ---------------------------------------------------------------------------------------------

/**
 * Characters never equal (a replaced word is replaced whole); a block equal to one of its type; an
 * atom equal to one of its type whose attributes differ in `fmt` alone, which a command setting its
 * formatting rewrites as a replace of the node: the atom is not replaced, and its formatting change
 * goes unrecorded (section 12.5 item 7). Once an atom was never equal, so bold over a tab put the
 * tab back deleted beside a bold one inserted (section 12.16).
 */
const ENCODER: TokenEncoder<unknown> = {
  encodeCharacter: () => ({}),
  encodeNodeStart: (node) => (node.isLeaf ? atomToken(node) : node.type.name),
  encodeNodeEnd: (node) => `/${node.type.name}`,
  compareTokens: (a, b) => a === b,
};

/**
 * The tokens for a transaction: the tracker's, with the end of each paragraph whose mark the
 * transaction names as new (`NEW_MARK_META`) equal to nothing. At a paragraph's end, Enter leaves
 * its old mark and a new one side by side with nothing between them; as tokens they are alike, and
 * the changeset took the second, the empty paragraph's, as the one inserted. Word inserts the
 * first, at the caret (core-ts check 29; ED-005 section 12.17).
 */
function encoderFor(tr: Transaction): TokenEncoder<unknown> {
  const made = new Set<PMNode>();
  for (const pos of newMarksOf(tr)) { const node = tr.doc.nodeAt(pos); if (node?.type.name === 'paragraph') made.add(node); }
  if (made.size === 0) return ENCODER;
  return { ...ENCODER, encodeNodeEnd: (node) => (made.has(node) ? {} : `/${node.type.name}`) };
}

/** The positions, in the transaction's document, of the paragraphs whose mark it made (`NEW_MARK_META`). */
function newMarksOf(tr: Transaction): number[] {
  return (tr.getMeta(NEW_MARK_META) as number[] | undefined) ?? [];
}

/** An atom as the changeset compares it: its type and its attributes but its formatting. */
function atomToken(node: PMNode): string {
  const { fmt: _fmt, ...rest } = node.attrs;
  void _fmt;
  return `${node.type.name} ${JSON.stringify(rest)}`;
}

const FORMATTING = new Set<string>(FORMATTING_MARKS);

/** The paragraph attributes that are its `w:pPr` properties, whose change is a `w:pPrChange`. */
const PARAGRAPH_PROPERTIES = ['pStyle', 'jc', 'ind', 'spacing', 'bidi', 'numPr', 'tabs'] as const;

/**
 * The transaction as tracked: `tr` with steps appended (section 5.4). Inserted content takes
 * the insertion mark (an atom, a paragraph's mark and a row the attribute), losing a deletion or
 * a formatting change it inherited. Deleted content is put back before what replaced it, with the
 * deletion mark, except what this author inserted, which goes for real, and what was deleted
 * already, which keeps its deletion. A paragraph mark deleted is kept, marked, instead of
 * joining. A formatting mark added or removed records the formatting before it, once; a change
 * back to it removes the record, as does a paragraph's properties changed back.
 */
export function trackTransaction(tr: Transaction, state: EditorState, tracking: Tracking, ctx: ProjectionContext): Transaction {
  const info: RevisionInfo = { author: tracking.author, date: tracking.date, move: null };
  const changes = ChangeSet.create(state.doc, undefined, encoderFor(tr)).addSteps(tr.doc, tr.mapping.maps, null).changes;
  const inserted = changes.filter((c) => c.toB > c.fromB).map((c) => [c.fromB, c.toB] as const);
  const insideInserted = (from: number, to: number): boolean => inserted.some(([a, b]) => a <= from && to <= b);
  // Marks and attributes first: they move no position, so the changes' positions hold after them.
  trackFormatting(tr, state, tracking, ctx, insideInserted);
  trackParagraphProperties(tr, state, tracking, ctx, insideInserted);
  const before = state.selection;
  let caret: number | undefined;
  const made: { from: number; to: number; at: number }[] = [];
  const deletedMarks: { pos: number; at: number }[] = [];
  for (const c of [...changes].sort((a, b) => b.fromB - a.fromB)) {
    if (c.toB > c.fromB) { markInserted(tr, c.fromB, c.toB, info); made.push({ from: c.fromB, to: c.toB, at: tr.mapping.maps.length }); }
    if (c.toA > c.fromA) {
      const size = restoreDeleted(tr, state, c.fromA, c.toA, c.fromB);
      markDeleted(tr, c.fromB, c.fromB + size, info, deletedMarks);
      // Backspace (what it deleted lay before the caret) leaves the caret before it, as Word does;
      // Delete, and a selection deleted, leave it after.
      if (changes.length === 1 && c.toB === c.fromB && before.empty && before.head >= c.toA) caret = c.fromB;
    }
  }
  afterItsDeletion(tr, made, info);
  carryDeletedMarksProperties(tr, deletedMarks, tracking, ctx);
  if (caret !== undefined) tr.setSelection(Selection.near(tr.doc.resolve(Math.min(caret, tr.doc.content.size)), 1));
  return tr;
}

/**
 * What replaced deleted text goes after all of it, as Word shows a replacement: the deletion, then
 * the insertion. A replace that keeps a zero-width item where it was (Word's `_GoBack` bookmark
 * inside a hit: `find.replace` moves no atom) leaves the transaction's deletion in two pieces with
 * the insertion after the first; the insertion moves past the rest (ED-005 section 12.7, found by
 * the agreement suite against the engine's tracked `insertText`).
 */
function afterItsDeletion(tr: Transaction, made: { from: number; to: number; at: number }[], info: RevisionInfo): void {
  const ours = (node: PMNode | null | undefined, name: 'insertion' | 'deletion'): boolean =>
    !!node?.isText && node.marks.some((m) => m.type.name === name && m.attrs['author'] === info.author && m.attrs['date'] === info.date);
  const ranges = made.map((m) => ({ from: tr.mapping.slice(m.at).map(m.from, 1), to: tr.mapping.slice(m.at).map(m.to, -1) })).sort((a, b) => b.from - a.from);
  for (const r of ranges) {
    const $from = tr.doc.resolve(r.from);
    if (!ours($from.nodeBefore, 'deletion')) continue;
    // Past zero-width items to the end of the deletion that goes on after the insertion.
    let end = r.to;
    let deletedAfter = false;
    for (;;) {
      const next = tr.doc.resolve(end).nodeAfter;
      if (!next) break;
      if (ours(next, 'deletion')) { deletedAfter = true; end += next.nodeSize; continue; }
      if (next.isInline && next.isLeaf && (next.type.name === 'opaqueInline') && ZERO_WIDTH.has(next.attrs['kind'] as string)) { end += next.nodeSize; continue; }
      break;
    }
    if (!deletedAfter) continue;
    // Back to the last deleted piece, so zero-width items after the deletion stay after the insertion.
    while (end > r.to && !ours(tr.doc.resolve(end).nodeBefore, 'deletion')) end -= tr.doc.resolve(end).nodeBefore!.nodeSize;
    const slice = tr.doc.slice(r.from, r.to);
    tr.insert(end, slice.content);
    tr.delete(r.from, r.to);
  }
}

/** Inline items with no content of their own, which a deletion runs on past. */
const ZERO_WIDTH = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'lastRenderedPageBreak']);

/** The inline, paragraph and row nodes a range holds, as the tracker marks them. */
interface Parts {
  text: [number, number, PMNode][];
  atoms: [number, PMNode][];
  /** Paragraphs whose mark (their closing token) is inside the range. */
  marks: [number, PMNode][];
  /** Rows wholly inside the range. */
  rows: [number, PMNode][];
}

function partsOf(doc: PMNode, from: number, to: number): Parts {
  const out: Parts = { text: [], atoms: [], marks: [], rows: [] };
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) { out.text.push([Math.max(pos, from), Math.min(pos + node.nodeSize, to), node]); return false; }
    if (node.isInline) { if (node.type.spec.attrs?.['revision'] !== undefined && pos >= from && pos + node.nodeSize <= to) out.atoms.push([pos, node]); return true; }
    if (node.type.name === 'paragraph') { const close = pos + node.nodeSize - 1; if (close >= from && close < to) out.marks.push([pos, node]); }
    if (node.type.name === 'table_row' && pos >= from && pos + node.nodeSize <= to) out.rows.push([pos, node]);
    return true;
  });
  return out;
}

/** Whether an insertion is this author's own, which a deletion removes for real: not a move's (core-ts check 20). */
function ownInsertion(r: RevisionInfo | undefined, author: string | null): boolean {
  return r !== undefined && r.author === author && (r.move ?? null) === null;
}

/** Inserted content: the insertion mark, less an inherited deletion or formatting change; atoms, paragraph marks and rows by attribute. */
function markInserted(tr: Transaction, from: number, to: number, info: RevisionInfo): void {
  const schema = tr.doc.type.schema;
  const insertion = schema.marks['insertion']!.create(info);
  const parts = partsOf(tr.doc, from, to);
  for (const [a, b] of parts.text) {
    tr.removeMark(a, b, schema.marks['deletion']!);
    tr.removeMark(a, b, schema.marks['formatChange']!);
    tr.addMark(a, b, insertion);
  }
  const revision: Revision = { insertion: info };
  for (const [pos] of parts.atoms) tr.setNodeAttribute(pos, 'revision', revision);
  for (const [pos] of parts.marks) tr.setNodeAttribute(pos, 'markRevision', revision);
  for (const [pos] of parts.rows) tr.setNodeAttribute(pos, 'revision', revision);
}

/** Put the deleted range of the old document back at `at`; the size put back. */
function restoreDeleted(tr: Transaction, state: EditorState, fromA: number, toA: number, at: number): number {
  const slice = state.doc.slice(fromA, toA);
  const before = tr.doc.content.size;
  tr.replace(at, at, slice);
  return tr.doc.content.size - before;
}

/**
 * Deleted content, put back: the deletion mark, or for what this author inserted, a deletion for
 * real (text and atoms removed, a paragraph mark joined, a row removed); what is deleted already
 * keeps its deletion. Moved-to text is not an insertion of its author's own: deleted, even by the
 * move's author, it keeps a deletion nested in the move, as Word writes it (core-ts check 20).
 */
function markDeleted(tr: Transaction, from: number, to: number, info: RevisionInfo, marked?: { pos: number; at: number }[]): void {
  const schema = tr.doc.type.schema;
  const deletion = schema.marks['deletion']!.create(info);
  const own = (r: RevisionInfo | undefined): boolean => ownInsertion(r, info.author);
  const parts = partsOf(tr.doc, from, to);
  const real: { from: number; to: number; join?: boolean }[] = [];
  for (const [a, b, node] of parts.text) {
    if (node.marks.some((m) => m.type.name === 'deletion')) continue;
    const ins = node.marks.find((m) => m.type.name === 'insertion');
    if (ins && own(ins.attrs as RevisionInfo)) real.push({ from: a, to: b });
    else tr.addMark(a, b, deletion);
  }
  for (const [pos, node] of parts.atoms) {
    const r = (node.attrs['revision'] as Revision | null) ?? {};
    if (r.deletion) continue;
    if (own(r.insertion)) real.push({ from: pos, to: pos + node.nodeSize });
    else tr.setNodeAttribute(pos, 'revision', { ...r, deletion: info });
  }
  for (const [pos, node] of parts.marks) {
    const r = (node.attrs['markRevision'] as Revision | null) ?? {};
    if (r.deletion) continue;
    if (own(r.insertion)) real.push({ from: pos + node.nodeSize, to: pos + node.nodeSize, join: true });
    else { tr.setNodeAttribute(pos, 'markRevision', { ...r, deletion: info }); marked?.push({ pos, at: tr.mapping.maps.length }); }
  }
  for (const [pos, node] of parts.rows) {
    const r = (node.attrs['revision'] as Revision | null) ?? {};
    if (r.deletion) continue;
    if (own(r.insertion)) real.push({ from: pos, to: pos + node.nodeSize });
    else tr.setNodeAttribute(pos, 'revision', { ...r, deletion: info });
  }
  // From the end, so each position still holds when it is used.
  for (const r of real.sort((a, b) => b.from - a.from || b.to - a.to)) {
    if (r.join) { if (canJoin(tr.doc, r.from)) tr.join(r.from); }
    else tr.delete(r.from, r.to);
  }
}

/**
 * A range marked deleted as the tracker marks what a transaction deleted, for a command that
 * deletes by marking: rows, whose deletion changes the table's structure only when it is
 * accepted, so a vertical merge through them stays whole (ED-005 section 12.5).
 */
export function markDeletedRange(tr: Transaction, from: number, to: number, tracking: Tracking): void {
  markDeleted(tr, from, to, { author: tracking.author, date: tracking.date, move: null });
}

/** A range marked inserted as the tracker marks what a transaction made, for a command that tracks by marking: a bound value (ED-005 section 12.57). */
export function markInsertedRange(tr: Transaction, from: number, to: number, tracking: Tracking): void {
  markInserted(tr, from, to, { author: tracking.author, date: tracking.date, move: null });
}

/**
 * A formatting mark added or removed (bold, a font, a character style): the text's formatting
 * before, as `w:rPrChange` holds it, in a `formatChange` mark, unless the text has one already
 * (a second change keeps the first `before`) or is this author's insertion; a change back to
 * `before`, compared by value, removes the mark (section 12.3 item 12).
 */
function trackFormatting(tr: Transaction, state: EditorState, tracking: Tracking, ctx: ProjectionContext, insideInserted: (from: number, to: number) => boolean): void {
  const ranges: [number, number][] = [];
  tr.steps.forEach((step, i) => {
    if (!(step instanceof AddMarkStep || step instanceof RemoveMarkStep) || !FORMATTING.has(step.mark.type.name)) return;
    const rest = tr.mapping.slice(i + 1);
    const from = rest.map(step.from, 1);
    const to = rest.map(step.to, -1);
    if (to > from) ranges.push([from, to]);
  });
  if (ranges.length === 0) return;
  const schema = tr.doc.type.schema;
  const formatChange = schema.marks['formatChange']!;
  const back = tr.mapping.invert();
  const writes: { from: number; to: number; add?: Mark }[] = [];
  for (const [from, to] of ranges) {
    state.doc.nodesBetween(back.map(from, 1), back.map(to, -1), (old, pos0) => {
      if (!old.isText) return true;
      const a0 = Math.max(pos0, back.map(from, 1));
      const b0 = Math.min(pos0 + old.nodeSize, back.map(to, -1));
      const a = tr.mapping.map(a0, 1);
      const b = tr.mapping.map(b0, -1);
      if (b <= a || insideInserted(a, b)) return false;
      tr.doc.nodesBetween(a, b, (now, pos) => {
        if (!now.isText) return true;
        const c = Math.max(pos, a);
        const d = Math.min(pos + now.nodeSize, b);
        const ins = now.marks.find((m) => m.type.name === 'insertion');
        if (ins && ownInsertion(ins.attrs as RevisionInfo, tracking.author)) return false;
        const after = runProperties(now.marks, ctx);
        const existing = now.marks.find((m) => m.type === formatChange);
        if (existing) {
          if (sameValue(JSON.parse(existing.attrs['before'] as string), JSON.parse(after))) writes.push({ from: c, to: d });
        } else {
          const beforeRPr = runProperties(old.marks, ctx);
          if (!sameValue(JSON.parse(beforeRPr), JSON.parse(after))) writes.push({ from: c, to: d, add: formatChange.create({ author: tracking.author, date: tracking.date, before: beforeRPr }) });
        }
        return false;
      });
      return false;
    });
  }
  for (const w of writes) {
    if (w.add) tr.addMark(w.from, w.to, w.add); else tr.removeMark(w.from, w.to, formatChange);
  }
}

/** A run's properties as a `formatChange` records them, from its text's marks and its source run's unseen properties. */
function runProperties(marks: readonly Mark[], ctx: ProjectionContext): string {
  const runSource = marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
  const source = runSource ? (ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
  const rPr = runPropertiesFor(source?.value.rPr, marks);
  if (rPr) delete rPr.rPrChange;
  return formatChangeBeforeRPr(rPr);
}

/** Whether a paragraph holds anything not deleted: text, or an atom, without a deletion. */
function keepsContent(paragraph: PMNode): boolean {
  let kept = false;
  paragraph.descendants((node) => {
    if (kept) return false;
    if (node.isText) { if (!node.marks.some((m) => m.type.name === 'deletion')) kept = true; return false; }
    if (node.isLeaf && !(node.attrs['revision'] as Revision | null | undefined)?.deletion) kept = true;
    return !node.isLeaf;
  });
  return kept;
}

/**
 * A paragraph mark deleted gives the next paragraph the deleted mark's paragraph's properties, with
 * a `w:pPrChange` recording its own, as Word writes it: Delete at the end of a heading with tracking
 * on leaves the heading as it was, its mark deleted, and the next paragraph a heading too, so that
 * the joined line is the heading's throughout once accepted (Jason's Word 15 measurement, ED-005
 * section 12.33 item 3; the Delete counterpart of core-ts check 29's Enter, section 12.17; the
 * engine's `Paragraph.delete()` from CR-002 section 38). Nothing when the two already agree. The
 * A paragraph deleted whole, its text and its mark, hands nothing on: Word carries nothing then (core-ts
 * check 33, ED-005 section 12.51); nor does an empty one, unmeasured. The
 * marks are taken in document order, so a run of deleted marks carries the first paragraph's
 * properties to the paragraph that ends it. The next paragraph's own record, if it has one, is kept
 * (a second change keeps the first `before`); one this author inserted takes the properties and no
 * record, as `trackParagraphProperties` treats it. A mark before a table or at a cell's end has no
 * next paragraph and changes nothing (unmeasured in Word).
 */
function carryDeletedMarksProperties(tr: Transaction, marked: readonly { pos: number; at: number }[], tracking: Tracking, ctx: ProjectionContext): void {
  const positions = marked.map((m) => tr.mapping.slice(m.at).map(m.pos, 1)).sort((a, b) => a - b);
  for (const pos of positions) {
    const deleted = tr.doc.nodeAt(pos);
    if (!deleted || deleted.type.name !== 'paragraph') continue;
    // A whole paragraph deleted, its text and its mark, hands nothing on (core-ts check 33, ED-005 section 12.51):
    // the carry is for a mark deleted alone, its paragraph's text kept to join the next.
    if (!keepsContent(deleted)) continue;
    const nextPos = pos + deleted.nodeSize;
    const next = tr.doc.resolve(nextPos).nodeAfter;
    if (!next || next.type.name !== 'paragraph') continue;
    if (PARAGRAPH_PROPERTIES.every((k) => plainEqual(deleted.attrs[k], next.attrs[k]))) continue;
    const attrs = next.attrs as unknown as ParagraphAttrs;
    const carried = Object.fromEntries(PARAGRAPH_PROPERTIES.map((k) => [k, deleted.attrs[k]]));
    const record = attrs.pPrChange ?? (ownInsertion(attrs.markRevision?.insertion, tracking.author)
      ? null
      : { author: tracking.author, date: tracking.date, before: propertiesBase(paragraphPropertiesOf(next, ctx)) });
    tr.setNodeMarkup(nextPos, undefined, { ...next.attrs, ...carried, pPrChange: record });
  }
}

/**
 * A paragraph's properties changed (a style, an alignment, an indent, a list): its `w:pPr` before
 * in `pPrChange`, unless it has one already or its mark is this author's insertion; changed back,
 * the record goes.
 */
function trackParagraphProperties(tr: Transaction, state: EditorState, tracking: Tracking, ctx: ProjectionContext, insideInserted: (from: number, to: number) => boolean): void {
  const positions = new Set<number>();
  tr.steps.forEach((step, i) => {
    let pos: number | undefined;
    if (step instanceof AttrStep) pos = step.pos;
    else if (step instanceof ReplaceAroundStep && (step as unknown as { structure: boolean }).structure && step.insert === 1 && step.gapFrom === step.from + 1 && step.gapTo === step.to - 1) pos = step.from;
    if (pos !== undefined) positions.add(tr.mapping.slice(i + 1).map(pos, 1));
  });
  if (positions.size === 0) return;
  const back = tr.mapping.invert();
  // The paragraph after a mark a split made keeps the old mark, and is the paragraph the split was
  // in, whatever its position maps back to (at the end, the text's end: core-ts check 29).
  const kept = new Map<number, number>();
  const madeMarks = new Set(newMarksOf(tr));
  for (const at of madeMarks) { const first = tr.doc.nodeAt(at); if (first) kept.set(at + first.nodeSize, back.map(at, 1)); }
  for (const pos of positions) {
    const now = tr.doc.nodeAt(pos);
    if (!now || now.type.name !== 'paragraph' || insideInserted(pos, pos + now.nodeSize)) continue;
    const old = state.doc.nodeAt(kept.get(pos) ?? back.map(pos, 1));
    if (!old || old.type.name !== 'paragraph') continue;
    const attrs = now.attrs as unknown as ParagraphAttrs;
    // A join moves the second paragraph's section onto the first; tracked, both paragraphs stay,
    // so the first keeps its own (a section change is never tracked: proposal 5). A split's new
    // paragraph has none of its own: the section stays with the old mark, after it (once the
    // section was put back on both halves, a section break more in the file: section 12.17).
    if (!madeMarks.has(pos) && !plainEqual(old.attrs['sectPr'], now.attrs['sectPr'])) tr.setNodeAttribute(pos, 'sectPr', old.attrs['sectPr']);
    if (ownInsertion(attrs.markRevision?.insertion, tracking.author)) continue;
    if (PARAGRAPH_PROPERTIES.every((k) => plainEqual(old.attrs[k], now.attrs[k]))) continue;
    const afterPPr = paragraphPropertiesOf(now, ctx);
    const after = propertiesBase(afterPPr);
    if (attrs.pPrChange) {
      if (sameValue(attrs.pPrChange.before, after)) tr.setNodeAttribute(pos, 'pPrChange', null);
      else if (backAsRecorded(attrs.pPrChange.before, after, afterPPr, state)) {
        // Back as it lays out, but written otherwise (Ctrl+L's `w:jc="left"`, an indent stepped back to 0):
        // the recorded properties go back and the record goes, as Word writes it (core-ts check 26).
        const recorded = projectedParagraphAttrs(revive(attrs.pPrChange.before) as wml.PPr | undefined, 'content') as unknown as Record<string, unknown>;
        tr.setNodeMarkup(pos, undefined, { ...now.attrs, ...Object.fromEntries(PARAGRAPH_PROPERTIES.map((k) => [k, recorded[k] ?? null])), pPrChange: null });
      }
    } else {
      const before = propertiesBase(paragraphPropertiesOf(old, ctx));
      if (!sameValue(before, after)) tr.setNodeAttribute(pos, 'pPrChange', { author: tracking.author, date: tracking.date, before });
    }
  }
}

/** The laid-out paragraph formatting, which Word compares when properties are changed back; everything else must be written the same. */
const LAYOUT = ['jc', 'ind', 'spacing'] as const;

const ALIGNED: Record<string, string> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'both', distribute: 'both' };

/**
 * Whether a paragraph's properties are back to the recorded ones as Word judges it (core-ts check
 * 26): the alignment, indents and spacing resolved through the styles equal, with Word's defaults
 * where nothing states one, and every other property written the same. Word's Ctrl+L removes a
 * `w:jc` the style gives and Decrease Indent a `w:ind` of 0, where the commands write the value.
 */
function backAsRecorded(before: unknown, after: unknown, afterPPr: wml.PPr | undefined, state: EditorState): boolean {
  const resolver = sessionOf(state)?.pkg.propertyResolverOrUndefined;
  if (!resolver) return false;
  const rest = (p: unknown): Record<string, unknown> => { const c = { ...(p as Record<string, unknown>) }; for (const k of LAYOUT) delete c[k]; return c; };
  if (!sameValue(rest(before), rest(after))) return false;
  const laidOut = (pPr: wml.PPr | undefined): string => {
    const e = paragraphAttrsOf(resolver.getEffectivePPr(pPr));
    const firstLine = e.ind?.hanging != null ? -e.ind.hanging : e.ind?.firstLine ?? 0;
    return JSON.stringify([ALIGNED[e.jc ?? 'left'] ?? e.jc, e.ind?.start ?? 0, e.ind?.end ?? 0, firstLine,
      e.spacing?.before ?? 0, e.spacing?.after ?? 0, e.spacing?.line ?? 240, e.spacing?.lineRule ?? 'auto']);
  };
  return laidOut(revive(before) as wml.PPr | undefined) === laidOut(afterPPr);
}

/** A `w:pPr` as `w:pPrChange` holds it: a `CT_PPrBase`, without the mark's properties, the section or a change of its own. */
function propertiesBase(pPr: wml.PPr | undefined): unknown {
  const plain = (pPr ? toPlain(pPr) : {}) as Record<string, unknown>;
  for (const k of ['rPr', 'sectPr', 'pPrChange', 'PARENT']) delete plain[k];
  plain['TYPE_NAME'] = 'org_docx4j_wml.PPrBase';
  return plain;
}

// ---------------------------------------------------------------------------------------------
// Tracking off
// ---------------------------------------------------------------------------------------------

/**
 * With tracking off, what a transaction makes is plain, as Word writes it (core-ts check 21,
 * measured for text): text and atoms typed inside or at the end of an insertion, or inside a
 * deletion, shed the revision they inherited, so the revision splits around them; so do a
 * paragraph mark and a row made there. Enter inside a paragraph whose mark is inserted gives the
 * new mark a copy of the paragraph's attributes, the revision included; the mark is plain, and the
 * other keeps the revision (Jason's finding, section 12.16). What the transaction made is what the
 * tracker would mark inserted with tracking on (the changeset over the tracker's tokens), so the
 * two agree on which mark Enter made, and a node whose attributes a command set (a paragraph's
 * alignment, an atom's formatting) is not made anew: it keeps its revision (ED-005 section 12.5
 * item 11).
 */
export function withoutInheritedRevisions(tr: Transaction): Transaction {
  const schema = tr.doc.type.schema;
  const changes = ChangeSet.create(tr.before, undefined, encoderFor(tr)).addSteps(tr.doc, tr.mapping.maps, null).changes;
  for (const c of changes) {
    if (c.toB <= c.fromB) continue;
    const parts = partsOf(tr.doc, c.fromB, c.toB);
    for (const [a, b, node] of parts.text) {
      for (const name of ['insertion', 'deletion'] as const) if (schema.marks[name]!.isInSet(node.marks)) tr.removeMark(a, b, schema.marks[name]!);
    }
    for (const [pos, node] of parts.atoms) if (node.attrs['revision']) tr.setNodeAttribute(pos, 'revision', null);
    for (const [pos, node] of parts.marks) if (node.attrs['markRevision']) tr.setNodeAttribute(pos, 'markRevision', null);
    for (const [pos, node] of parts.rows) if (node.attrs['revision']) tr.setNodeAttribute(pos, 'revision', null);
  }
  return tr;
}

/** A range's deletions taken off: text, atoms, paragraph marks and rows inside it. */
function shedDeletion(tr: Transaction, from: number, to: number): void {
  const deletion: MarkType = tr.doc.type.schema.marks['deletion']!;
  const shed = (r: Revision | null): Revision | null => { if (!r?.deletion) return r; const { deletion: _d, ...rest } = r; void _d; return rest.insertion ? rest : null; };
  const parts = partsOf(tr.doc, from, to);
  for (const [a, b, node] of parts.text) if (deletion.isInSet(node.marks)) tr.removeMark(a, b, deletion);
  for (const [pos, node] of parts.atoms) if ((node.attrs['revision'] as Revision | null)?.deletion) tr.setNodeAttribute(pos, 'revision', shed(node.attrs['revision'] as Revision));
  for (const [pos, node] of parts.marks) if ((node.attrs['markRevision'] as Revision | null)?.deletion) tr.setNodeAttribute(pos, 'markRevision', shed(node.attrs['markRevision'] as Revision));
  for (const [pos, node] of parts.rows) if ((node.attrs['revision'] as Revision | null)?.deletion) tr.setNodeAttribute(pos, 'revision', shed(node.attrs['revision'] as Revision));
}

// ---------------------------------------------------------------------------------------------
// An IME composition (section 5.4)
// ---------------------------------------------------------------------------------------------

/**
 * A composition's text, once it is over, recorded as a transaction's would have been: the view
 * applies a composition's transactions as they come, since marking text while it is composed
 * breaks the composition in Chrome, and keeps the document as it was when the composition began
 * and every step map since. ProseMirror replaces the whole composed text node as the composition
 * goes, so the maps overstate what changed; the changes here are the documents compared
 * character by character (the changeset's own tokens), so only what the composition added is
 * inserted and only what it replaced is put back, deleted. With tracking off, a deletion the
 * composed text inherited is shed.
 */
export function markComposed(tr: Transaction, start: PMNode, maps: readonly StepMap[], tracking: Tracking | undefined): Transaction {
  const changes = ChangeSet.create(start).addSteps(tr.doc, maps, null).changes;
  if (!tracking) {
    for (const c of changes) if (c.toB > c.fromB) shedDeletion(tr, c.fromB, c.toB);
    return tr;
  }
  const info: RevisionInfo = { author: tracking.author, date: tracking.date, move: null };
  for (const c of [...changes].sort((a, b) => b.fromB - a.fromB)) {
    if (c.toB > c.fromB) markInserted(tr, c.fromB, c.toB, info);
    if (c.toA > c.fromA) {
      const slice = start.slice(c.fromA, c.toA);
      const before = tr.doc.content.size;
      tr.replace(c.fromB, c.fromB, slice);
      markDeleted(tr, c.fromB, c.fromB + tr.doc.content.size - before, info);
    }
  }
  return tr;
}

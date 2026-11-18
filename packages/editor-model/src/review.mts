// SPDX-License-Identifier: GPL-3.0-or-later
//
// Accept and reject (ED-005 sections 5.5 and 5.6, step C4): the tracked changes the edit layer
// holds, listed as Office JS lists them, and resolved as ProseMirror transforms over the marks
// and attributes, each one transaction and so one undo step. A change is what Office JS counts as
// one (core-ts checks 22 and 23): pieces that touch, of one type (inserted or deleted, a paragraph
// mark included, its text "\r") and by one author, whatever their dates. A table row inserted or
// deleted is one change with its cells' revisions in it, as the engine reports it. A move's half
// is one change and is resolved with its other half, as Word resolves it (check 18); an edited or a
// broken move as checks 20 measured. The revisions the edit layer only preserves (a section's, a
// table's, a cell's properties changed) are not among them: the engine lists none of them either.
// `preservedChangesOf` lists them apart, for the review panel to show as Word's (section 6.5).
import type { Mark, Node as PMNode, Schema } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import * as Y from 'yjs';
import { initProseMirrorDoc, absolutePositionToRelativePosition, relativePositionToAbsolutePosition, ySyncPluginKey } from 'y-prosemirror';
import { canJoin } from 'prosemirror-transform';
import { TableMap, removeRow } from 'prosemirror-tables';
import { rPrFromElements } from '@docx4j/generated-objects-ts/builders/wml';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { readRunOptions } from '@docx4j/core-ts';
import { FORMATTING_MARKS, type IndentAttrs, type ParagraphAttrs, type Revision, type RevisionInfo } from './schema.mjs';
import { formattingMarks, projectedParagraphAttrs } from './import.mjs';
import { revive } from './json.mjs';
import { dateString } from './revisions.mjs';

/** Office JS `Word.ChangeTrackingState`, as a change lists it. */
export type ChangeType = 'Added' | 'Deleted' | 'Formatted';

type Side = 'insertion' | 'deletion' | 'format' | 'properties';

/** One piece of revision markup: text, an atom, a paragraph mark, a row, a run's or a paragraph's properties. */
export interface ChangePiece {
  what: 'text' | 'atom' | 'mark' | 'row' | 'format' | 'properties';
  side: Side;
  author: string | null;
  date: string | null;
  /** The move a piece belongs to (the range's name, `''` when it has none); null for any other. */
  move: string | null;
  /** The content it covers: text's range, an atom's or a row's node, a paragraph mark's closing token. */
  from: number;
  to: number;
  /** The node a mark, an atom, a row or a paragraph's properties are on. */
  pos: number;
  /** The mark instance on text (an insertion, a deletion, a formatting change). */
  mark?: Mark;
  /** A formatting change's `before`, a paragraph's properties' change. */
  before?: unknown;
  text: string;
  /**
   * The run holders it is in: a link, an inline or block control, a simple field (`''` for none). Two
   * pieces join only in one holder, as Office JS lists them (core-ts check 30: an insertion that runs
   * into a link and out again is three changes). Keys are positions within the piece's top-level block.
   */
  holder?: string;
  /** A row's own pieces: its cells' revisions of the row's side, which it lists and resolves with it. */
  inner?: ChangePiece[];
  /** A revision around a run with nothing in it (`isEmptyRun`): it rides with a change it touches, and is no change of its own. */
  empty?: true;
}

/** A tracked change as Office JS lists it: its type, author, first piece's date, text, and where it is. */
export interface TrackedChangeInfo {
  type: ChangeType;
  author: string | null;
  date: string | null;
  text: string;
  from: number;
  to: number;
  /** A move's half: its name. */
  move?: string;
  /** Rows inserted or deleted, and nothing else: every piece a row (core-ts check 31: rows group with what touches them). */
  row?: boolean;
  /** Only revisions around empty runs: never listed, resolved by "all" alone (`emptyChangesOf`). */
  empty?: true;
  pieces: ChangePiece[];
}

/**
 * Items with no content of their own, which do not keep two revisions from touching. A comment's
 * reference run among them: Office JS lists two insertions either side of a plain one as one change
 * (Word check 32, ED-005 section 12.49).
 */
const MARKERS = new Set(['bookmarkStart', 'bookmarkEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'permStart', 'permEnd', 'lastRenderedPageBreak', 'proofErr', 'r']);

const isMarker = (node: PMNode): boolean => (node.type.name === 'opaqueInline' || node.type.name === 'opaqueBlock') && MARKERS.has(node.attrs['kind'] as string);

/**
 * A run with nothing in it, the atom a `w:r` of no items projects to. A revision around one marks nothing a
 * reader could see or decide on: ONLYOFFICE leaves three around each heading's bookmark when it refreshes a
 * Contents with tracking on (ED-005 section 12.113), so it draws no change bar and lists no change.
 */
export const isEmptyRun = (node: PMNode): boolean => node.type.name === 'opaqueInline' && node.attrs['kind'] === 'r';

/** A revision's author as the listing has it: `w:author=""` and no `w:author` at all are one unknown author, `null`. */
const authorOf = (author: string | null | undefined): string | null => author || null;

/** The key of the link a text node is in (`''` for none), as a piece's holder writes it. */
function linkKey(node: PMNode): string {
  const link = node.marks.find((m) => m.type.name === 'link');
  return link ? `link:${String(link.attrs['sourceId'])}|${String(link.attrs['rId'])}|${String(link.attrs['anchor'])};` : '';
}

const runOf = (node: PMNode): unknown => (node.isText ? node.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] : node.attrs['runSource']);

/**
 * The link each atom among a paragraph's inline children is in, by child index. An atom carries no marks
 * (UPSTREAM-2), so the text around it says: text of one link on both sides, or the link's text in the atom's own
 * run. An atom at a link's edge in a run of its own says nothing and counts as outside it.
 */
function atomLinks(parent: PMNode): Map<number, string> {
  const out = new Map<number, string>();
  const children: PMNode[] = [];
  parent.forEach((child) => children.push(child));
  if (!children.some((child) => child.isText && linkKey(child) !== '')) return out;
  const textAfter: (PMNode | undefined)[] = [];
  let next: PMNode | undefined;
  for (let i = children.length - 1; i >= 0; i--) {
    textAfter[i] = next;
    if (children[i]!.isText) next = children[i];
    else if (!children[i]!.isLeaf) next = undefined;
  }
  let before: PMNode | undefined;
  children.forEach((child, i) => {
    if (child.isText) { before = child; return; }
    if (!child.isLeaf) { before = undefined; return; }
    const after = textAfter[i];
    const a = before ? linkKey(before) : '';
    const b = after ? linkKey(after) : '';
    const run = runOf(child);
    const sharesRun = (text: PMNode | undefined): boolean => run !== undefined && text !== undefined && runOf(text) === run;
    const key = a !== '' && (a === b || sharesRun(before)) ? a : b !== '' && sharesRun(after) ? b : '';
    if (key !== '') out.set(i, key);
  });
  return out;
}

/** The text Office JS gives an atom. */
function atomText(node: PMNode): string {
  if (node.type.name === 'tab') return '\t';
  if (node.type.name === 'hardBreak') return node.attrs['kind'] === 'page' ? '\f' : '\v';
  return '';
}

const sidesOf = (r: Revision | null | undefined): ('insertion' | 'deletion')[] => [...(r?.insertion ? ['insertion' as const] : []), ...(r?.deletion ? ['deletion' as const] : [])];

/**
 * A row's text as Office JS reports it: its cells with a tab between them, then `\r\n`; a cell's
 * paragraphs with `\r` between them, a nested table's rows each with their own `\r\n`, a block
 * control's paragraphs as if it were not there (core-ts checks 30 and 31, the engine's `rowText`).
 * An inserted row's text as accepted, a deleted row's as it was.
 */
export function rowTextOf(row: PMNode, side: 'insertion' | 'deletion'): string {
  const skip = side === 'insertion' ? 'deletion' : 'insertion';
  const paragraphText = (p: PMNode): string => {
    let s = '';
    p.descendants((n) => {
      if (n.isText) { if (!n.marks.some((m) => m.type.name === skip)) s += n.text; return false; }
      if (n.isLeaf) { if (!(n.attrs['revision'] as Revision | null | undefined)?.[skip]) s += atomText(n); return false; }
      return true;
    });
    return s;
  };
  const cellText = (cell: PMNode): string => {
    let out = '';
    let afterParagraph = false;
    const visit = (node: PMNode): void => {
      if (node.type.name === 'paragraph') { out += (afterParagraph ? '\r' : '') + paragraphText(node); afterParagraph = true; }
      else if (node.type.name === 'table') { node.forEach((r) => { if (r.type.name === 'table_row') out += rowTextOf(r, side); }); afterParagraph = false; }
      else node.forEach(visit);
    };
    cell.forEach(visit);
    return out;
  };
  const cells: string[] = [];
  row.forEach((cell) => cells.push(cellText(cell)));
  return `${cells.join('\t')}\r\n`;
}

/** Every piece of revision markup in a document, in document order (a paragraph's mark after its text, a row as one piece). */
function piecesOf(doc: PMNode): ChangePiece[] {
  const pieces: ChangePiece[] = [];
  const inline = (node: PMNode, at: number, into: ChangePiece[], holder: string): void => {
    if (node.isText) {
      const own = `${holder}${linkKey(node)}`;
      for (const m of node.marks) {
        const name = m.type.name;
        if (name !== 'insertion' && name !== 'deletion' && name !== 'formatChange') continue;
        const side: Side = name === 'formatChange' ? 'format' : name;
        into.push({ what: side === 'format' ? 'format' : 'text', side, author: authorOf(m.attrs['author'] as string | null), date: (m.attrs['date'] as string | null) ?? null, move: side === 'format' ? null : ((m.attrs['move'] as string | null) ?? null), from: at, to: at + node.nodeSize, pos: at, mark: m, ...(side === 'format' ? { before: m.attrs['before'] } : {}), text: node.text ?? '', holder: own });
      }
      return;
    }
    const own = node.type.name === 'opaqueInline' && node.attrs['kind'] === 'fldSimple' ? `${holder}fld@${at};` : holder;
    const r = node.attrs['revision'] as Revision | null | undefined;
    for (const side of sidesOf(r)) {
      const info = r![side]!;
      into.push({ what: 'atom', side, author: authorOf(info.author), date: info.date, move: info.move, from: at, to: at + node.nodeSize, pos: at, text: atomText(node), holder: own, ...(isEmptyRun(node) ? { empty: true as const } : {}) });
    }
  };
  /** A paragraph's pieces: its properties' change, its inline content (an inline control's inside with a holder of its own), its mark with `markHolder`. */
  const paragraph = (node: PMNode, pos: number, into: ChangePiece[], holder = '', markHolder = holder): void => {
    const attrs = node.attrs as unknown as ParagraphAttrs;
    if (attrs.pPrChange) into.push({ what: 'properties', side: 'properties', author: authorOf(attrs.pPrChange.author), date: attrs.pPrChange.date, move: null, from: pos, to: pos + node.nodeSize, pos, before: attrs.pPrChange.before, text: '', holder });
    const walk = (parent: PMNode, base: number, h: string): void => {
      const links = atomLinks(parent);
      parent.forEach((child, offset, index) => {
        const at = base + offset;
        inline(child, at, into, `${h}${links.get(index) ?? ''}`);
        if (child.type.name === 'sdtInline') walk(child, at + 1, `${h}sdt@${at};`);
      });
    };
    walk(node, pos + 1, holder);
    const close = pos + node.nodeSize - 1;
    for (const side of sidesOf(attrs.markRevision)) {
      const info = attrs.markRevision![side]!;
      into.push({ what: 'mark', side, author: authorOf(info.author), date: info.date, move: info.move, from: close, to: close + 1, pos, text: '\r', holder: markHolder });
    }
  };
  /**
   * The blocks under `parent`, with the holder their pieces are in: a block control's inside has one
   * of its own, but its last paragraph's mark is the outside's, as Office JS lists it (core-ts check 30).
   */
  const blocks = (parent: PMNode, base: number, holder: string): void => parent.forEach((node, offset) => {
    const pos = base + offset;
    if (node.type.name === 'sdtBlock') {
      const inside = `${holder}sdtb@${pos};`;
      node.forEach((child, childOffset, index) => {
        const at = pos + 1 + childOffset;
        if (child.type.name === 'paragraph') paragraph(child, at, pieces, inside, index === node.childCount - 1 ? holder : inside);
        else blocks(child, at + 1, inside);
      });
      return;
    }
    if (node.type.name === 'table_row') {
      const r = node.attrs['revision'] as Revision | null;
      const side = r?.insertion ? 'insertion' : r?.deletion ? 'deletion' : undefined;
      // A plain row keeps its cells' changes to itself: they group among themselves, never with what
      // lies outside the row (the engine's break around a plain row, core-ts check 31).
      if (!side) { blocks(node, pos + 1, `${holder}row@${pos};`); return; }
      // A row inserted or deleted is one piece, its cells' revisions of its side in it, grouping with the
      // rows and the text that touch it as any piece does: neither a row's nor a table's edge parts them
      // (core-ts check 31, Jason's decision 2026-10-01).
      const inner: ChangePiece[] = [];
      node.descendants((child, rel) => { if (child.type.name === 'paragraph') paragraph(child, pos + 1 + rel, inner); return child.type.name !== 'paragraph'; });
      const info = r![side]!;
      pieces.push({ what: 'row', side, author: authorOf(info.author), date: info.date, move: null, from: pos, to: pos + node.nodeSize, pos, text: rowTextOf(node, side), holder, inner: inner.filter((p) => p.side === side) });
      return;
    }
    if (node.type.name === 'paragraph') { paragraph(node, pos, pieces, holder); return; }
    blocks(node, pos + 1, holder);
  });
  blocks(doc, 0, '');
  return pieces;
}

/** Whether nothing with content lies between two positions: no text, no content atom, no block's end. */
function touches(doc: PMNode, a: number, b: number): boolean {
  if (b <= a) return true;
  let ok = true;
  doc.nodesBetween(a, b, (node, pos) => {
    if (!ok) return false;
    const end = pos + node.nodeSize;
    if (node.isText) { if (pos < b && end > a) ok = false; return false; }
    if (node.isLeaf) { if (pos >= a && pos < b && !isMarker(node)) ok = false; return false; }
    // A block control's end is no content: its holder parts what is inside from what is after (core-ts
    // check 30). Nor is a table's: a row that is a revision groups with what follows the table (check 31).
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'table' && end - 1 >= a && end - 1 < b) { ok = false; return false; }
    return true;
  });
  return ok;
}

const typeOf = (side: Side): ChangeType => (side === 'insertion' ? 'Added' : side === 'deletion' ? 'Deleted' : 'Formatted');

/** Whether a piece joins the grouping's chain of its side: a named move's half and a paragraph's properties never do. */
const chained = (p: ChangePiece): boolean => p.side !== 'properties' && (p.move === null || p.move === '');

/**
 * One top-level block's revision markup, positions relative to the block (the block at 0), and
 * what the grouping needs to join its pieces with other blocks' without walking the document
 * between them (ED-005 section 6.5): whether two pieces touch is decided by what lies between
 * them, and between two blocks that is the end of the first, any blocks in the middle, and the
 * start of the second.
 */
interface Listing {
  /** The changes, as `trackedChangesOf` gives them. */
  changes: TrackedChangeInfo[];
  /** The revisions around empty runs that touch no change, as `emptyChangesOf` gives them. */
  empties: TrackedChangeInfo[];
}

interface BlockListing {
  pieces: ChangePiece[];
  /** Per piece: whether it touches the previous piece of its side in this block; undefined for the first of its side here. */
  touchesPrevious: (boolean | undefined)[];
  /** Per piece: nothing with content between the block's start and it (read for the first of its side). */
  clearBefore: boolean[];
  /** Per piece: nothing with content between it and the block's end (read for the last of its side). */
  clearAfter: boolean[];
  /** A marker between blocks (a bookmark, a move's range), which two pieces touch across. */
  marker: boolean;
}

const blockListings = new WeakMap<PMNode, BlockListing>();
const documentListings = new WeakMap<PMNode, Listing>();
let blockWalks = 0;

/** How many blocks the listing has walked since the module loaded (a test's spy: a keystroke should walk one). */
export function listingBlockWalks(): number {
  return blockWalks;
}

/** A block alone in a document of its own, at position 0, so the document-level functions read it with positions relative to it. */
function alone(block: PMNode): PMNode {
  return block.type.schema.topNodeType.create(null, block);
}

function blockListing(block: PMNode): BlockListing {
  let out = blockListings.get(block);
  if (out) return out;
  blockWalks++;
  const doc = alone(block);
  const pieces = piecesOf(doc);
  const touchesPrevious: (boolean | undefined)[] = [];
  const clearBefore: boolean[] = [];
  const clearAfter: boolean[] = [];
  const previous = new Map<Side, ChangePiece>();
  for (const p of pieces) {
    const prev = chained(p) ? previous.get(p.side) : undefined;
    touchesPrevious.push(prev ? touches(doc, prev.to, p.from) && (prev.holder ?? '') === (p.holder ?? '') : undefined);
    clearBefore.push(touches(doc, 0, p.from));
    clearAfter.push(touches(doc, p.to, block.nodeSize));
    if (chained(p) && !p.empty) previous.set(p.side, p);
  }
  out = { pieces, touchesPrevious, clearBefore, clearAfter, marker: isMarker(block) };
  blockListings.set(block, out);
  return out;
}

const shiftPiece = (p: ChangePiece, by: number): ChangePiece => ({ ...p, from: p.from + by, to: p.to + by, pos: p.pos + by, ...(p.inner ? { inner: p.inner.map((q) => shiftPiece(q, by)) } : {}) });

/**
 * The tracked changes of a document, in document order, as Office JS lists them (ED-005 section
 * 5.5): touching pieces of one type by one author are one change, a row inserted or deleted among
 * them (core-ts check 31); a move's half is one change whatever it spans (never grouped with
 * others); a paragraph's properties changed is one change. A nested revision (another author's
 * insertion deleted) is listed apart.
 *
 * Incremental (ED-005 section 6.5): each top-level block's pieces are memoised per block node,
 * and the grouping is one pass over the blocks' lists, so after a keystroke only the block it
 * touched is walked; the list is memoised per document node. The result is a fresh array of
 * shared change objects, which a caller must not change.
 *
 * A revision around an empty run is no change (ED-005 section 12.113): it joins a change it touches, to
 * be resolved with it, and otherwise is left to `emptyChangesOf`. Neither way does it part a change.
 */
export function trackedChangesOf(doc: PMNode): TrackedChangeInfo[] {
  return listingOf(doc).changes.slice();
}

/** The revisions around empty runs that touch no change: not listed, and resolved by accept all and reject all alone, so that those leave none in the file. */
export function emptyChangesOf(doc: PMNode): TrackedChangeInfo[] {
  return listingOf(doc).empties.slice();
}

function listingOf(doc: PMNode): Listing {
  const known = documentListings.get(doc);
  if (known) return known;
  const changes: TrackedChangeInfo[] = [];
  const empties: TrackedChangeInfo[] = [];
  /** Per side, the change the next piece may join, and whether nothing with content lies between its last piece and the block being read (for a last piece in an earlier block). */
  const open = new Map<Side, { change: TrackedChangeInfo; clear: boolean; holder: string }>();
  const moves = new Map<string, TrackedChangeInfo>();
  const start = (p: ChangePiece): TrackedChangeInfo => ({ type: typeOf(p.side), author: p.author, date: p.date, text: '', from: p.from, to: p.to, pieces: [], ...(p.what === 'row' ? { row: true } : {}) });
  const add = (c: TrackedChangeInfo, p: ChangePiece): void => {
    c.pieces.push(p, ...(p.inner ?? []));
    c.text += p.text;
    c.to = Math.max(c.to, p.to);
    if (p.what !== 'row' && c.row) delete c.row;
  };
  doc.forEach((block, offset) => {
    const listing = blockListing(block);
    /** The index of each side's last chained piece in this block. */
    const last = new Map<Side, number>();
    listing.pieces.forEach((relative, i) => {
      const p = shiftPiece(relative, offset);
      if (p.side === 'properties') { const c = start(p); add(c, p); changes.push(c); return; }
      if (!chained(p)) {
        const key = `${p.side}:${p.move}`;
        let c = moves.get(key);
        if (!c) { c = { ...start(p), move: p.move! }; moves.set(key, c); changes.push(c); }
        add(c, p);
        return;
      }
      const o = open.get(p.side);
      // Across top-level blocks only outside every holder: a holder never spans two (its keys are positions within one).
      const touching = last.has(p.side) ? listing.touchesPrevious[i]! : (o?.clear ?? false) && listing.clearBefore[i]! && o!.holder === '' && (p.holder ?? '') === '';
      const joins = o !== undefined && o.change.author === p.author && touching;
      if (p.empty) {
        if (joins) add(o.change, p);
        else { const c: TrackedChangeInfo = { ...start(p), empty: true }; add(c, p); empties.push(c); }
        return;
      }
      if (joins) { add(o.change, p); o.holder = p.holder ?? ''; }
      else { const c = start(p); add(c, p); changes.push(c); open.set(p.side, { change: c, clear: false, holder: p.holder ?? '' }); }
      last.set(p.side, i);
    });
    // What lies between each side's last piece and the next block: the rest of this block, or this whole block.
    for (const [side, o] of open) {
      const i = last.get(side);
      o.clear = i !== undefined ? listing.clearAfter[i]! : o.clear && listing.marker;
    }
  });
  const order: Record<ChangeType, number> = { Added: 0, Deleted: 1, Formatted: 2 };
  const out: Listing = { changes: changes.sort((a, b) => a.from - b.from || order[a.type] - order[b.type]), empties };
  documentListings.set(doc, out);
  return out;
}

// ---------------------------------------------------------------------------------------------
// The revisions shown in Word only (ED-005 sections 5.2 and 6.5)
// ---------------------------------------------------------------------------------------------

/** The revisions the edit layer keeps as they are, by element (the diagnostics' `tracked.preserved` keys). */
export type PreservedKind = 'sectPrChange' | 'tblPrChange' | 'tblGridChange' | 'trPrChange' | 'tcPrChange' | 'cellIns' | 'cellDel' | 'cellMerge' | 'numberingChange' | 'markRPrChange'
  // A revision inside an opaque atom (core-ts 0.6.0; ED-008 section 8.8 items 10 and 14 D3): an equation's text or
  // structure inserted or deleted in Word's form (`w:ins`/`w:del` in its math runs and `m:ctrlPr`), and a content
  // control or custom XML element taken whole under tracking (`w:customXmlInsRangeStart`/`DelRangeStart` pairs
  // around it). The projection shows neither as a change; the engine lists and resolves both.
  | 'mathIns' | 'mathDel' | 'holderIns' | 'holderDel';

/** The first revision inside an equation's object: a math run's `w:ins`/`w:del` item or a structure's `m:ctrlPr` record. */
function equationRevisionOf(value: unknown): { kind: 'ins' | 'del'; record: Recorded } | undefined {
  let found: { kind: 'ins' | 'del'; record: Recorded } | undefined;
  const visit = (o: unknown, depth: number): void => {
    if (found || o === null || typeof o !== 'object' || depth > 64) return;
    if (Array.isArray(o)) { for (const x of o) visit(x, depth + 1); return; }
    const obj = o as Record<string, unknown>;
    const name = obj['name'] as { localPart?: string } | undefined;
    if (name && (name.localPart === 'ins' || name.localPart === 'del') && obj['value'] && typeof obj['value'] === 'object') {
      found = { kind: name.localPart, record: obj['value'] as Recorded };
      return;
    }
    for (const key of ['ins', 'del'] as const) {
      const v = obj[key] as { author?: string } | undefined;
      if (v && typeof v === 'object' && 'author' in v) { found = { kind: key, record: v as Recorded }; return; }
    }
    for (const key of Object.keys(obj)) if (key !== 'PARENT') visit(obj[key], depth + 1);
  };
  visit(value, 0);
  return found;
}

/** The property records of a table, which Office JS lists as one change per table (core-ts check 27, CR-002 section 35). */
const TABLE_RECORDS = new Set<PreservedKind>(['tblPrChange', 'tblGridChange', 'trPrChange', 'tcPrChange']);

/** One revision shown in Word only: its kind, who made it and when, and the node it is on. */
export interface PreservedChange {
  kind: PreservedKind;
  author: string | null;
  date: string | null;
  /** The node it is on (a paragraph, a table, a row, a cell); the document's end for the final section's. */
  pos: number;
}

/** The value of the element a node came from, by `sourceId`, for the two records no attribute holds (`tblGridChange`, `numberingChange`). */
export type SourceLookup = (sourceId: string) => unknown;

type Recorded = { author?: string; date?: wml.XmlCalendar } | null | undefined;

const preservedListings = new WeakMap<PMNode, PreservedChange[]>();

function preservedIn(block: PMNode, source: SourceLookup): PreservedChange[] {
  let out = preservedListings.get(block);
  if (out) return out;
  out = [];
  const found = out;
  const note = (kind: PreservedKind, value: Recorded, pos: number, fallback?: Recorded): void => {
    if (!value) return;
    const who = value.author !== undefined || value.date !== undefined ? value : fallback ?? value;
    found.push({ kind, author: who?.author ?? null, date: dateString(who), pos });
  };
  const value = (node: PMNode): Record<string, unknown> | undefined => {
    const id = node.attrs['sourceId'] as string | null;
    return (id ? source(id) : undefined) as Record<string, unknown> | undefined;
  };
  // The revisions inside a paragraph's opaque atoms (core-ts 0.6.0): an equation's, and a control's taken whole.
  const atomsIn = (paragraph: PMNode, base: number): void => {
    paragraph.descendants((node, rel) => {
      if (node.type.name !== 'opaqueInline') return true;
      const kind = node.attrs['kind'] as string;
      if (kind === 'oMath' || kind === 'oMathPara') {
        const rev = equationRevisionOf(value(node));
        if (rev) note(rev.kind === 'ins' ? 'mathIns' : 'mathDel', rev.record, base + rel);
      } else if (kind === 'customXmlInsRangeStart' || kind === 'customXmlDelRangeStart') {
        // The outer pair's start, right before the element it takes; the inner pair's start ends the element's content.
        const next = paragraph.resolve(rel + node.nodeSize).nodeAfter;
        if (next && (next.type.name === 'sdtInline' || next.type.name === 'sdtBlock')) note(kind === 'customXmlInsRangeStart' ? 'holderIns' : 'holderDel', value(node) as Recorded, base + rel);
      }
      return false;
    });
  };
  alone(block).descendants((node, pos) => {
    const a = node.attrs as Record<string, Record<string, Recorded> | null>;
    switch (node.type.name) {
      case 'paragraph':
        note('sectPrChange', a['sectPr']?.['sectPrChange'], pos);
        note('markRPrChange', a['markRPr']?.['rPrChange'], pos);
        note('numberingChange', ((value(node)?.['pPr'] as { numPr?: { numberingChange?: Recorded } } | undefined)?.numPr?.numberingChange), pos);
        atomsIn(node, pos + 1);
        return false;
      case 'table': {
        const tblPrChange = a['tblPr']?.['tblPrChange'];
        note('tblPrChange', tblPrChange, pos);
        // A grid's change carries no author or date of its own: it is the table's change's.
        note('tblGridChange', (value(node)?.['tblGrid'] as { tblGridChange?: Recorded } | undefined)?.tblGridChange, pos, tblPrChange);
        return true;
      }
      case 'table_row':
        note('trPrChange', a['trPr']?.['trPrChange'], pos);
        return true;
      case 'table_cell':
        for (const kind of ['tcPrChange', 'cellIns', 'cellDel', 'cellMerge'] as const) note(kind, a['tcPr']?.[kind], pos);
        return true;
      default:
        return true;
    }
  });
  preservedListings.set(block, out);
  return out;
}

/**
 * The revisions shown in Word only as Office JS lists them (core-ts check 27, CR-002 section 35):
 * a table's property records (its own, its grid's, its rows' and its cells') are one change per
 * table, at the table, as Word writes a set of them for one property change; the others as they
 * are (a cell inserted, deleted or merged, a section's properties, a list's numbering, a mark's
 * formatting). The review panel's entries; the diagnostics count by record.
 */
export interface PreservedGroup {
  kind: PreservedKind | 'tableProperties';
  author: string | null;
  date: string | null;
  pos: number;
  /** The records it holds: one, or a table's several. */
  records: PreservedChange[];
}

export function preservedGroupsOf(doc: PMNode, records: readonly PreservedChange[]): PreservedGroup[] {
  const out: PreservedGroup[] = [];
  const tables = new Map<number, PreservedGroup>();
  for (const r of records) {
    if (!TABLE_RECORDS.has(r.kind)) { out.push({ kind: r.kind, author: r.author, date: r.date, pos: r.pos, records: [r] }); continue; }
    // The table the record is on, or the one its row or cell is in.
    const $pos = doc.resolve(r.pos);
    let table = r.pos;
    if (doc.nodeAt(r.pos)?.type.name !== 'table') for (let d = $pos.depth; d >= 1; d--) if ($pos.node(d).type.name === 'table') { table = $pos.before(d); break; }
    let group = tables.get(table);
    if (!group) { group = { kind: 'tableProperties', author: r.author, date: r.date, pos: table, records: [] }; tables.set(table, group); out.push(group); }
    group.records.push(r);
  }
  return out.sort((a, b) => a.pos - b.pos);
}

/**
 * The revisions a document keeps and shows in Word only (ED-005 section 5.2's preserved list), in
 * document order, for the review panel (section 6.5): a section's, a table's, a grid's, a row's
 * and a cell's properties changed, a cell inserted, deleted or merged, a list's numbering, a
 * paragraph mark's formatting. The body's only: a story keeps E1's projection (section 6.9).
 * Memoised per top-level block, as `trackedChangesOf` is.
 */
export function preservedChangesOf(doc: PMNode, source: SourceLookup): PreservedChange[] {
  const out: PreservedChange[] = [];
  doc.forEach((block, offset) => { for (const p of preservedIn(block, source)) out.push({ ...p, pos: p.pos + offset }); });
  const final = (doc.attrs['finalSectPr'] as { sectPrChange?: Recorded } | null)?.sectPrChange;
  if (final) out.push({ kind: 'sectPrChange', author: final.author ?? null, date: dateString(final), pos: doc.content.size });
  return out;
}

// ---------------------------------------------------------------------------------------------
// Resolving
// ---------------------------------------------------------------------------------------------

/** The element a projected item stands for, by `sourceId`: a move's range markers name and number it. */
export type ElementLookup = (sourceId: string) => { name?: unknown; id?: unknown } | undefined;

/** Whether every attribute an indent states is the level's own value, so the level gives it all. */
function givenByLevel(ind: IndentAttrs, level: IndentAttrs | null): boolean {
  if (!level) return false;
  const stated = (Object.keys(ind) as (keyof IndentAttrs)[]).filter((k) => ind[k] !== null && ind[k] !== undefined);
  return stated.length > 0 && stated.every((k) => ind[k] === level[k]);
}

/** What resolving reads from the document beyond the edit layer. */
export interface ResolveContext {
  element: ElementLookup;
  /** A list level's own indent, for a numbering change rejected (`definitions.indOf`). */
  listIndent?: (numId: string, ilvl: number) => IndentAttrs | null;
}

/** An edit that moves positions, applied last, from the end: a range deleted, a paragraph mark joined, a row removed. */
type Shift = { at: number; run: (tr: Transaction) => void };

/**
 * Accept or reject one change in `tr` (ED-005 section 5.5): an insertion accepted and a deletion
 * rejected lose their mark; an insertion rejected and a deletion accepted go, text and atoms
 * removed, a paragraph mark joining its paragraph with the next, which keeps its own properties,
 * a row removed. A formatting change accepted loses its mark, rejected restores `before`; a
 * paragraph's properties likewise. A move's half resolves the whole move (section 5.6).
 */
export function resolveChange(tr: Transaction, change: TrackedChangeInfo, action: 'accept' | 'reject', context: ResolveContext): void {
  if (change.move !== undefined) { resolveMove(tr, change, action, context); return; }
  const shifts: Shift[] = [];
  piecesToResolve(change, action).forEach((p) => resolvePiece(tr, p, action, shifts, context));
  applyShifts(tr, shifts);
  // A move half with no range markers is resolved as the plain change it is: what is left of its move goes plain.
  if (change.pieces.some((p) => p.move === '')) dissolveBrokenMoves(tr);
}

/**
 * What is left of a broken move once part of it is resolved becomes a plain pending insertion or
 * deletion, its range markers kept, as Word leaves it (check 20, case 13; core-ts fix E): a half
 * with no range markers, and a named half whose other half is gone.
 */
function dissolveBrokenMoves(tr: Transaction): void {
  const pieces = piecesOf(tr.doc);
  const sides = new Map<string, Set<Side>>();
  for (const p of pieces) if (p.move) (sides.get(p.move) ?? sides.set(p.move, new Set()).get(p.move)!).add(p.side);
  const broken = (p: ChangePiece): boolean => p.move === '' || (p.move !== null && sides.get(p.move)!.size < 2);
  for (const p of pieces) {
    if (!broken(p) || (p.side !== 'insertion' && p.side !== 'deletion')) continue;
    if (p.what === 'text') {
      tr.removeMark(p.from, p.to, p.mark!);
      tr.addMark(p.from, p.to, p.mark!.type.create({ ...p.mark!.attrs, move: null }));
    } else {
      const attr = p.what === 'mark' ? 'markRevision' : 'revision';
      const r = { ...(tr.doc.nodeAt(p.pos)!.attrs[attr] as Revision) };
      r[p.side] = { ...r[p.side]!, move: null };
      tr.setNodeAttribute(p.pos, attr, r);
    }
  }
}

/**
 * A row that goes goes whole: its cells' pieces are resolved only when it stays (a cell emptied on
 * the way would be invalid). A change may hold several rows and text around them (core-ts check 31),
 * so each row decides for its own cells.
 */
function piecesToResolve(change: TrackedChangeInfo, action: 'accept' | 'reject'): ChangePiece[] {
  const skipped = new Set<ChangePiece>();
  for (const p of change.pieces) if (p.what === 'row' && (p.side === 'insertion') !== (action === 'accept')) for (const q of p.inner ?? []) skipped.add(q);
  return skipped.size === 0 ? change.pieces : change.pieces.filter((p) => !skipped.has(p));
}

/**
 * Several changes resolved in one transaction, from the last to the first: each is found again in
 * the document as it is by then (its start mapped, its type and author), so a move, which edits two
 * places, cannot leave another change's positions stale; a change already resolved with a move's
 * other half is skipped. The number resolved, a revision around an empty run being no change.
 */
export function resolveChanges(tr: Transaction, changes: readonly TrackedChangeInfo[], action: 'accept' | 'reject', context: ResolveContext): number {
  const start = tr.mapping.maps.length;
  const targets = [...changes].sort((a, b) => b.from - a.from);
  const done = new Set<string>();
  let count = 0;
  for (const target of targets) {
    if (target.move !== undefined && done.has(target.move)) continue;
    const from = tr.mapping.slice(start).map(target.from, 1);
    const same = (c: TrackedChangeInfo): boolean => c.from === from && c.type === target.type && c.author === target.author && c.move === target.move;
    const listing = listingOf(tr.doc);
    const now = listing.changes.find(same) ?? listing.empties.find(same);
    if (!now) continue;
    resolveChange(tr, now, action, context);
    if (now.move !== undefined) done.add(now.move);
    if (!now.empty) count++;
  }
  return count;
}

/** y-prosemirror's map between the store's types and a document's nodes. */
type ProsemirrorMapping = Parameters<typeof absolutePositionToRelativePosition>[2];

/**
 * The live map of a view's state (y-prosemirror's binding, which the sync plugin keeps as the
 * view's document changes), when the state has one and it is over this fragment (ED-005 section
 * 6.5): a key made through it costs a walk to the position, where a map rebuilt from the store
 * costs the whole document. Headless (the runner's states, the tests) there is none.
 */
export function liveMapping(state: EditorState | undefined, fragment: Y.XmlFragment): ProsemirrorMapping | undefined {
  const sync = state ? (ySyncPluginKey.getState(state) as { type?: unknown; binding?: { mapping: ProsemirrorMapping } | null } | undefined) : undefined;
  return sync?.type === fragment && sync.binding ? sync.binding.mapping : undefined;
}

/**
 * A change's key (ED-005 section 5.5): its start as a Yjs relative position in the store, which
 * survives remote edits around it, with its type, author and move, since a revision nested in
 * another (a deletion inside moved or inserted text) starts where the other does. Through a view's
 * live map when one is given (`liveMapping`), else a map rebuilt from the store.
 */
export function changeKey(fragment: Y.XmlFragment, schema: Schema, change: Pick<TrackedChangeInfo, 'from' | 'type' | 'author' | 'move'>, live?: ProsemirrorMapping): string {
  const mapping = live ?? initProseMirrorDoc(fragment, schema).mapping;
  const at = Y.relativePositionToJSON(absolutePositionToRelativePosition(change.from, fragment, mapping)) as unknown;
  return JSON.stringify({ at, type: change.type, author: change.author, ...(change.move !== undefined ? { move: change.move } : {}) });
}

/** The change a key names in a document now, or undefined when it is gone (resolved, or its text deleted by a peer); through a view's live map when one is given. */
export function changeOfKey(fragment: Y.XmlFragment, schema: Schema, doc: PMNode, key: string, live?: ProsemirrorMapping): TrackedChangeInfo | undefined {
  const k = JSON.parse(key) as { at: object; type: ChangeType; author: string | null; move?: string };
  const mapping = live ?? initProseMirrorDoc(fragment, schema).mapping;
  const pos = relativePositionToAbsolutePosition(fragment.doc!, fragment, Y.createRelativePositionFromJSON(k.at), mapping) as number | null;
  if (pos === null) return undefined;
  return trackedChangesOf(doc).find((c) => c.from === pos && c.type === k.type && c.author === k.author && c.move === k.move);
}

/**
 * A replacement's pairs (ED-005 proposal 41, section 9.7): each deletion with the insertion right
 * beside it, by the same author, with nothing that has content between them (the listing's own
 * `touches`, which skips markers). So across a paragraph's end only when the deletion took that
 * paragraph's mark: a kept mark is content between them. Either order (ED-008 section 8.8 item 9):
 * a person typing over a selection, Word and the editor's recorder write the deletion first; a
 * script's `insertText(…, 'Replace')` through Office JS and the engine since core-ts 0.6.0 writes
 * the insertion first (check 54 case 4), and the document holds whichever its maker wrote. A move's
 * half and a row never pair; dates are not compared, by choice. Nothing is stored: the review panel
 * pairs at each listing, and the listing itself stays in Office JS's form. Deletion to insertion.
 */
export function replacementPairs(doc: PMNode, changes: readonly TrackedChangeInfo[]): Map<TrackedChangeInfo, TrackedChangeInfo> {
  const out = new Map<TrackedChangeInfo, TrackedChangeInfo>();
  const taken = new Set<TrackedChangeInfo>();
  const plain = (c: TrackedChangeInfo): boolean => c.move === undefined && c.pieces.every((p) => p.move === null && p.what !== 'row');
  for (let i = 0; i < changes.length; i++) {
    const first = changes[i]!;
    if ((first.type !== 'Deleted' && first.type !== 'Added') || !plain(first) || taken.has(first)) continue;
    const other = first.type === 'Deleted' ? 'Added' : 'Deleted';
    for (let j = i + 1; j < changes.length; j++) {
      const c = changes[j]!;
      if (c.from < first.to) continue;
      if (!touches(doc, first.to, c.from)) break;
      if (c.type === other && plain(c) && c.author === first.author && !taken.has(c)) {
        if (first.type === 'Deleted') out.set(first, c); else out.set(c, first);
        taken.add(first);
        taken.add(c);
        break;
      }
    }
  }
  return out;
}

/**
 * What a formatting or properties change changed (ED-005 section 6.5, the review panel's words):
 * the formatting marks that differ between `before` and the text now (`font.*` per field), or the
 * paragraph's editable properties that differ (`ind.*` and `spacing.*` per field). Empty for any
 * other change.
 */
export function whatChanged(doc: PMNode, change: TrackedChangeInfo): string[] {
  const out = new Set<string>();
  for (const p of change.pieces) for (const k of changedIn(doc, p)) out.add(k);
  return [...out];
}

const sameValue = (a: unknown, b: unknown): boolean => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
const fieldOf = (x: unknown, f: string): unknown => (x as Record<string, unknown> | null)?.[f] ?? null;

/** What one formatting or properties piece changed: `whatChanged`'s keys. */
function changedIn(doc: PMNode, p: ChangePiece): string[] {
  const out: string[] = [];
  if (p.what === 'format') {
    const node = doc.nodeAt(p.from);
    if (!node) return out;
    const plain = JSON.parse(p.before as string) as unknown;
    const rPr = rPrFromElements(plain === null ? undefined : (revive(plain) as wml.CTRPrChange['rPr']));
    const before = formattingMarks(readRunOptions(rPr), rPr, doc.type.schema);
    for (const name of FORMATTING_MARKS) {
      const a = before.find((m) => m.type.name === name);
      const b = node.marks.find((m) => m.type.name === name);
      if (name === 'font') {
        for (const f of ['name', 'size', 'color', 'highlight']) if (!sameValue(a?.attrs[f], b?.attrs[f])) out.push(`font.${f}`);
      } else if (!a !== !b || (a && b && !a.eq(b))) out.push(name);
    }
  } else if (p.what === 'properties') {
    const node = doc.nodeAt(p.pos);
    if (!node) return out;
    const before = projectedParagraphAttrs((p.before ? revive(p.before) : undefined) as wml.PPr | undefined, 'content') as unknown as Record<string, unknown>;
    const now = node.attrs as Record<string, unknown>;
    for (const k of ['pStyle', 'jc', 'bidi', 'numPr', 'tabs']) if (!sameValue(before[k], now[k])) out.push(k);
    for (const f of ['start', 'end']) if (!sameValue(fieldOf(before['ind'], f), fieldOf(now['ind'], f))) out.push(`ind.${f}`);
    if (!sameValue(fieldOf(before['ind'], 'firstLine'), fieldOf(now['ind'], 'firstLine')) || !sameValue(fieldOf(before['ind'], 'hanging'), fieldOf(now['ind'], 'hanging'))) out.push('ind.firstLine');
    for (const f of ['before', 'after', 'line']) if (!sameValue(fieldOf(before['spacing'], f), fieldOf(now['spacing'], f))) out.push(`spacing.${f}`);
  }
  return out;
}

/** A property a formatting or properties change changed, and its value now; `null` where the change took the property away. */
export interface ChangedValue {
  /** `whatChanged`'s key: a formatting mark's name, `font.*`, or a paragraph property (`ind.*`, `spacing.*` per field). */
  property: string;
  /**
   * The value as the text or the paragraph has it now: `true` for a mark with no value of its own
   * (`strike`'s is `'double'` for a double one); a mark's `val` (underline, vertical alignment); a
   * character style's id; a font's name, size (points), colour or highlight (`'#RRGGBB'` or a highlight
   * name); a paragraph's style id, alignment, direction (a boolean), numbering (`{ numId, ilvl }`) or tab
   * stops; an indent or spacing in twips (a first-line indent negative for a hanging one); a line spacing
   * as `{ line, rule }`.
   */
  value: unknown;
}

/**
 * What a formatting or properties change changed, with the values it now has (ED-005 proposal 43, the
 * review panel's words as Word's Reviewing Pane has them: "Font: 16 pt", "Centred", "Heading 2"). The
 * application words them; this returns data, each property read from the first piece that changed it.
 */
export function changedValues(doc: PMNode, change: TrackedChangeInfo): ChangedValue[] {
  const out = new Map<string, ChangedValue>();
  for (const p of change.pieces) {
    for (const property of changedIn(doc, p)) {
      if (out.has(property)) continue;
      const node = doc.nodeAt(p.what === 'format' ? p.from : p.pos);
      out.set(property, { property, value: node ? (p.what === 'format' ? runValue(node, property) : paragraphValue(node, property)) : null });
    }
  }
  return [...out.values()];
}

function runValue(node: PMNode, property: string): unknown {
  const mark = (name: string): Mark | undefined => node.marks.find((m) => m.type.name === name);
  if (property.startsWith('font.')) return mark('font')?.attrs[property.slice('font.'.length)] ?? null;
  const m = mark(property);
  if (!m) return null;
  if (property === 'underline' || property === 'vertAlign') return m.attrs['val'] ?? true;
  if (property === 'strike') return m.attrs['double'] ? 'double' : true;
  if (property === 'runStyle') return m.attrs['id'] ?? null;
  return true;
}

function paragraphValue(node: PMNode, property: string): unknown {
  const attrs = node.attrs as Record<string, unknown>;
  if (property === 'ind.firstLine') {
    const hanging = fieldOf(attrs['ind'], 'hanging');
    return typeof hanging === 'number' && hanging !== 0 ? -hanging : fieldOf(attrs['ind'], 'firstLine');
  }
  if (property.startsWith('ind.')) return fieldOf(attrs['ind'], property.slice('ind.'.length));
  if (property === 'spacing.line') {
    const line = fieldOf(attrs['spacing'], 'line');
    return line === null ? null : { line, rule: fieldOf(attrs['spacing'], 'lineRule') };
  }
  if (property.startsWith('spacing.')) return fieldOf(attrs['spacing'], property.slice('spacing.'.length));
  return attrs[property] ?? null;
}

function resolvePiece(tr: Transaction, p: ChangePiece, action: 'accept' | 'reject', shifts: Shift[], context: ResolveContext): void {
  const keep = (p.side === 'insertion') === (action === 'accept');
  switch (p.what) {
    case 'text':
      if (keep) tr.removeMark(p.from, p.to, p.mark!);
      else shifts.push({ at: p.from, run: (t) => t.delete(p.from, p.to) });
      return;
    case 'atom':
      if (keep) dropSide(tr, p.pos, 'revision', p.side as 'insertion' | 'deletion');
      else shifts.push({ at: p.from, run: (t) => t.delete(p.from, p.to) });
      return;
    case 'mark':
      if (keep) dropSide(tr, p.pos, 'markRevision', p.side as 'insertion' | 'deletion');
      else shifts.push({ at: p.to, run: (t) => removeParagraphMark(t, p.pos) });
      return;
    case 'row':
      if (keep) dropSide(tr, p.pos, 'revision', p.side as 'insertion' | 'deletion');
      else shifts.push({ at: p.from, run: (t) => removeTableRow(t, p.pos) });
      return;
    case 'format': {
      if (action === 'reject') restoreFormatting(tr, p);
      tr.removeMark(p.from, p.to, p.mark!);
      return;
    }
    case 'properties': {
      const node = tr.doc.nodeAt(p.pos)!;
      const attrs: Record<string, unknown> = { ...node.attrs, pPrChange: null };
      if (action === 'reject') {
        const before = projectedParagraphAttrs((p.before ? revive(p.before) : undefined) as wml.PPr | undefined, 'content') as unknown as Record<string, unknown>;
        for (const k of ['pStyle', 'jc', 'ind', 'spacing', 'bidi', 'numPr', 'tabs']) attrs[k] = before[k] ?? null;
        // Word records a list item's indent from its level with the properties before a numbering
        // change, whole or in part, and puts the list back when the change is rejected, the indent
        // coming from the level again, not written on the paragraph: an indent every attribute of
        // which is the level's goes (check 18's reject-all.docx, section 9; check 25's 25b, where
        // hanging=360 alone was recorded against the level's left=720 hanging=360).
        const numPr = attrs['numPr'] as { numId: string | null; ilvl: number } | null;
        if (numPr?.numId && attrs['ind'] && context.listIndent && givenByLevel(attrs['ind'] as IndentAttrs, context.listIndent(numPr.numId, numPr.ilvl))) attrs['ind'] = null;
      }
      tr.setNodeMarkup(p.pos, undefined, attrs);
      return;
    }
  }
}

/** A revision's side taken off a node's attribute (`revision`, `markRevision`), the other side kept. */
function dropSide(tr: Transaction, pos: number, attr: 'revision' | 'markRevision', side: 'insertion' | 'deletion'): void {
  const r = { ...((tr.doc.nodeAt(pos)!.attrs[attr] as Revision | null) ?? {}) };
  delete r[side];
  tr.setNodeAttribute(pos, attr, r.insertion || r.deletion ? r : null);
}

/** A formatting change rejected: the run's formatting marks become what `before` says. */
function restoreFormatting(tr: Transaction, p: ChangePiece): void {
  const schema = tr.doc.type.schema;
  const plain = JSON.parse(p.before as string) as unknown;
  const rPr = rPrFromElements(plain === null ? undefined : (revive(plain) as wml.CTRPrChange['rPr']));
  for (const name of FORMATTING_MARKS) tr.removeMark(p.from, p.to, schema.marks[name]!);
  for (const m of formattingMarks(readRunOptions(rPr), rPr, schema)) tr.addMark(p.from, p.to, m);
}

/**
 * Apply the edits that move positions from the last to the first, so each position holds when it
 * is used; at one position a range marker goes first (`at` half a step on), so a paragraph mark
 * removed there joins the paragraph after it rather than the marker.
 */
function applyShifts(tr: Transaction, shifts: Shift[]): void {
  for (const s of shifts.sort((a, b) => b.at - a.at)) s.run(tr);
}

/**
 * A paragraph's mark removed (an insertion rejected, a deletion accepted): the paragraph joins the
 * next, whose mark survives and so whose properties the joined paragraph takes, as in Word. With
 * no paragraph after it to join, an empty paragraph goes whole, and one with text keeps its mark.
 */
function removeParagraphMark(tr: Transaction, pos: number): void {
  const node = tr.doc.nodeAt(pos);
  if (!node || node.type.name !== 'paragraph') return;
  const boundary = pos + node.nodeSize;
  const next = tr.doc.nodeAt(boundary);
  if (next && next.type === node.type && canJoin(tr.doc, boundary)) {
    tr.join(boundary);
    tr.setNodeMarkup(pos, undefined, next.attrs);
    return;
  }
  const $pos = tr.doc.resolve(pos);
  if (node.content.size === 0 && $pos.parent.childCount > 1) { tr.delete(pos, boundary); return; }
  if (node.attrs['markRevision']) tr.setNodeAttribute(pos, 'markRevision', null);
}

/** A row removed as prosemirror-tables removes one, a vertical merge through it shortened. */
function removeTableRow(tr: Transaction, rowPos: number): void {
  const $row = tr.doc.resolve(rowPos);
  const table = $row.parent;
  const tableStart = $row.start();
  const index = $row.index();
  if (table.childCount === 1) {
    // The last row goes with its table.
    const tablePos = tableStart - 1;
    tr.delete(tablePos, tablePos + table.nodeSize);
    return;
  }
  const map = TableMap.get(table);
  removeRow(tr, { map, table, tableStart, left: 0, top: index, right: map.width, bottom: index + 1 }, index);
}

// ---------------------------------------------------------------------------------------------
// Moves (section 5.6)
// ---------------------------------------------------------------------------------------------

/** A move's range markers in the document: the start and end of either side, found by the start's name and the end's id. */
function moveMarkers(doc: PMNode, name: string, lookup: ElementLookup): { from: number[]; to: number[]; toRange?: [number, number] } {
  const found: { kind: string; pos: number; size: number; id: string; name: string | undefined }[] = [];
  doc.descendants((node, pos) => {
    if (!isMarker(node)) return true;
    const kind = node.attrs['kind'] as string;
    if (!kind.startsWith('move')) return false;
    const el = lookup(node.attrs['sourceId'] as string);
    found.push({ kind, pos, size: node.nodeSize, id: String(el?.id ?? ''), name: typeof el?.name === 'string' ? el.name : undefined });
    return false;
  });
  const out: { from: number[]; to: number[]; toRange?: [number, number] } = { from: [], to: [] };
  for (const side of ['moveFrom', 'moveTo'] as const) {
    const start = found.find((f) => f.kind === `${side}RangeStart` && f.name === name);
    if (!start) continue;
    const end = found.find((f) => f.kind === `${side}RangeEnd` && f.id === start.id);
    const list = side === 'moveFrom' ? out.from : out.to;
    list.push(start.pos);
    if (end) list.push(end.pos);
    if (side === 'moveTo' && end) out.toRange = [start.pos, end.pos];
  }
  return out;
}

/**
 * A move resolved from either half (core-ts checks 18 and 20): accepted, the destination keeps its
 * text plain and the source goes, text and mark; rejected, the source keeps its text plain and
 * everything inside the destination's range goes, what was typed or deleted there since included,
 * as Word's "Keep original location text" does (text typed before the range start stays). Edits
 * inside moved text stay pending when the move is accepted. All four range markers go. A half
 * whose other half is missing resolves as the plain change it stands for; a destination left
 * alone keeps its text even when rejected (check 20, case 12).
 */
function resolveMove(tr: Transaction, change: TrackedChangeInfo, action: 'accept' | 'reject', context: ResolveContext): void {
  const name = change.move!;
  const all = trackedChangesOf(tr.doc).filter((c) => c.move === name);
  const dest = all.find((c) => c.type === 'Added');
  const source = all.find((c) => c.type === 'Deleted');
  const markers = moveMarkers(tr.doc, name, context.element);
  const shifts: Shift[] = [];
  if (!dest || !source) {
    // A broken move: the half there is resolves as the plain change it stands for, except that a
    // destination with no source keeps its text rejected as accepted (Word, case 12).
    const half = dest ?? source!;
    const plainAction = half === dest ? 'accept' : action;
    for (const p of half.pieces) resolvePiece(tr, p, plainAction, shifts, context);
    for (const pos of [...markers.from, ...markers.to]) shifts.push({ at: pos + 0.5, run: (t) => deleteNodeAt(t, pos) });
    applyShifts(tr, shifts);
    dissolveBrokenMoves(tr);
    return;
  }
  if (action === 'accept') {
    for (const p of dest.pieces) resolvePiece(tr, p, 'accept', shifts, context);
    for (const p of source.pieces) resolvePiece(tr, p, 'accept', shifts, context);
  } else {
    for (const p of source.pieces) resolvePiece(tr, p, 'reject', shifts, context);
    const [a, b] = markers.toRange ?? [dest.from, dest.to];
    removeDestination(tr, a, b, shifts);
  }
  for (const pos of [...markers.from, ...markers.to]) shifts.push({ at: pos + 0.5, run: (t) => deleteNodeAt(t, pos) });
  applyShifts(tr, shifts);
}

/** Everything between a destination's range markers: its text and atoms removed, the paragraph marks inside joined (the paragraphs go). */
function removeDestination(tr: Transaction, a: number, b: number, shifts: Shift[]): void {
  tr.doc.nodesBetween(a, b, (node, pos) => {
    if (node.type.name === 'paragraph') {
      const close = pos + node.nodeSize - 1;
      if (close > a && close < b) shifts.push({ at: close + 1, run: (t) => removeParagraphMark(t, pos) });
      return true;
    }
    if (node.isInline && !isMarker(node) && pos >= a && pos + node.nodeSize <= b) {
      shifts.push({ at: pos, run: (t) => t.delete(pos, pos + node.nodeSize) });
      return false;
    }
    return true;
  });
}

/** A range marker removed, inline or between blocks. */
function deleteNodeAt(tr: Transaction, pos: number): void {
  const node = tr.doc.nodeAt(pos);
  if (node && isMarker(node)) tr.delete(pos, pos + node.nodeSize);
}
